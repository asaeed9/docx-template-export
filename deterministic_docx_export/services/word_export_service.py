# Copyright © 2024 Ahsan Saeed
# Licensed under the Apache License, Version 2.0
# See LICENSE and NOTICE files for details.

"""
Word export service for deterministic document export.

This module provides functionality to export structured content (markdown or plain text)
into Word (.docx) templates. The export process is strictly non-generative: it does
not call any LLMs and does not modify wording. It only maps existing content into
Word document structures.

Key features:
- Scalar placeholder replacement (e.g., {{document_id}}, {{title}})
- Block placeholder replacement with structured content ({{summary}}, {{proposal}}, etc.)
- Markdown parsing and conversion to Word structures (headings, lists, tables)
- Manual list rendering: bullet and numbered lists use deterministic glyph/number insertion
  (no Word numbering XML) for stable, cross-platform rendering
- Plain text mode fallback (paragraphs only, no structure inference)
- Combined markdown file export (exported_markdown_content.md) for content analysis

List Rendering Approach:
- Bullet lists: Manual glyph insertion with configurable glyph sets and indentation
  policies defined by ListRenderConfig. Glyphs are inserted directly into paragraph text
  with manual indentation. Maximum visual depth is configurable via ListRenderConfig.
  Deep nesting behavior beyond max_visual_depth is deterministic and policy-driven
  (clamp_last, cycle, or textual strategies).
- Numbered lists: Manual hierarchical numbering (1., 1.1., 1.1.1.) with Python-based
  counter tracking per list block. Numbers reset for each new list block.
- Both list types use configurable manual paragraph indentation (default: 0.25" per level,
  -0.25" hanging indent, configurable via ListRenderConfig)

List Semantics:
- List rendering is block-based, not AST-linked. Each markdown list block is processed
  independently and rendered as a separate visual block in Word.
- Nested lists are preserved visually (indentation and glyphs reflect nesting depth),
  not structurally (no Word list object relationships).
- Mixed bullet ↔ numbered lists are rendered as separate blocks by design. This is a
  deliberate stability decision to ensure deterministic output and cross-platform
  consistency.

Table Behavior:
- Inline markdown formatting inside table cells is treated as literal text.
- Tables prioritize structure and layout determinism over inline formatting.

Deep Nesting Behavior:
- Very deep nesting (beyond max_visual_depth) may reduce readability due to Word layout
  constraints. This is a limitation of Word document layout, not the export engine.
- The engine handles deep nesting deterministically according to the configured strategy
  (clamp_last, cycle, or textual), but visual clarity may degrade with extreme nesting.

The module respects content fidelity: all text is preserved exactly as provided,
with only structural transformations applied (markdown syntax → Word objects).
"""
import logging
import os
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import List, Optional, Dict, Any, Iterable, Tuple, Union

from docx import Document  # type: ignore
from docx.document import Document as DocumentType  # type: ignore
from docx.text.paragraph import Paragraph  # type: ignore
from docx.shared import Inches  # type: ignore

from markdown_it import MarkdownIt  # type: ignore

from deterministic_docx_export.models.export_config import ListRenderConfig
from deterministic_docx_export.models.export_models import WordExportRequest
from deterministic_docx_export.services.output_path import get_output_dir_for_document, ensure_output_dir_exists

logger = logging.getLogger(__name__)


class Region(Enum):
    """Document region where content is located."""
    BODY = "BODY"
    HEADER = "HEADER"
    FOOTER = "FOOTER"


class Container(Enum):
    """Container type for content."""
    FLOW = "FLOW"  # Normal paragraphs
    TABLE = "TABLE"  # Table cells
    TEXTBOX = "TEXTBOX"  # Textboxes (w:txbxContent)


@dataclass
class FormattedRun:
    """Represents a text run with formatting."""
    text: str
    bold: bool = False
    italic: bool = False


@dataclass
class MarkdownBlock:
    """
    Simple normalized representation of a markdown block.

    For bullet lists and numbered lists, ``items`` contains a list of
    (level, text, formatted_runs) tuples, where level starts at 1 for
    top-level items, 2 for nested items, etc. This allows us to preserve
    nested list structure and formatting when exporting to Word.
    
    For headings and paragraphs, ``formatted_runs`` contains a list of
    FormattedRun objects that preserve bold/italic formatting. If None,
    falls back to plain ``text``.
    """

    type: str  # "heading" | "paragraph" | "bullet_list" | "numbered_list" | "table"
    text: Optional[str] = None  # for headings and paragraphs (fallback if formatted_runs not available)
    formatted_runs: Optional[List[FormattedRun]] = None  # for headings and paragraphs with formatting
    level: Optional[int] = None  # for headings
    items: Optional[List[tuple[int, str, Optional[List[FormattedRun]]]]] = None  # for bullet_list and numbered_list (level, text, formatted_runs)
    header: Optional[List[str]] = None  # for table header row
    rows: Optional[List[List[str]]] = None  # for table body rows


def _create_markdown_parser() -> MarkdownIt:
    """
    Create a markdown-it parser instance configured for commonmark with
    tables and lists enabled.
    """
    md = MarkdownIt("commonmark").enable("table").enable("list")
    return md


def _parse_inline_formatting(inline_token) -> List[FormattedRun]:
    """
    Parse inline markdown tokens to extract text with formatting (bold, italic).
    
    Returns a list of FormattedRun objects representing text segments with
    their formatting applied.
    """
    runs: List[FormattedRun] = []
    if not inline_token or not hasattr(inline_token, 'children'):
        return runs
    
    children = inline_token.children or []
    
    def parse_children(start_idx: int, bold_context: bool = False, italic_context: bool = False) -> int:
        """
        Recursively parse children tokens, tracking bold/italic context.
        Returns the index after processing.
        """
        idx = start_idx
        while idx < len(children):
            child = children[idx]
            
            if child.type == "text":
                if child.content:
                    runs.append(FormattedRun(
                        text=child.content,
                        bold=bold_context,
                        italic=italic_context
                    ))
                idx += 1
            
            elif child.type == "strong_open":
                # Enter bold context
                idx += 1
                # Check for nested em (bold-italic)
                if idx < len(children) and children[idx].type == "em_open":
                    idx += 1
                    idx = parse_children(idx, bold_context=True, italic_context=True)
                    if idx < len(children) and children[idx].type == "em_close":
                        idx += 1
                else:
                    idx = parse_children(idx, bold_context=True, italic_context=italic_context)
                
                # Skip strong_close
                if idx < len(children) and children[idx].type == "strong_close":
                    idx += 1
            
            elif child.type == "em_open":
                # Enter italic context
                idx += 1
                # Check for nested strong (bold-italic)
                if idx < len(children) and children[idx].type == "strong_open":
                    idx += 1
                    idx = parse_children(idx, bold_context=True, italic_context=True)
                    if idx < len(children) and children[idx].type == "strong_close":
                        idx += 1
                else:
                    idx = parse_children(idx, bold_context=bold_context, italic_context=True)
                
                # Skip em_close
                if idx < len(children) and children[idx].type == "em_close":
                    idx += 1
            
            elif child.type in ("strong_close", "em_close"):
                # Exit formatting context - return to caller
                return idx
            
            else:
                # Unknown token - skip
                idx += 1
        
        return idx
    
    parse_children(0)
    return runs


def parse_markdown_to_blocks(text: str) -> List[MarkdownBlock]:
    """
    Parse markdown text into a list of structural blocks.

    Blocks supported:
        - heading
        - paragraph
        - bullet_list
        - table

    This function MUST NOT change wording ? it only discovers structure.
    On any parsing error, it falls back to plain paragraphs split on
    blank lines.
    """
    if not text or not text.strip():
        return []

    try:
        md = _create_markdown_parser()
        tokens = md.parse(text)

        blocks: List[MarkdownBlock] = []

        i = 0
        while i < len(tokens):
            tok = tokens[i]

            # Headings
            if tok.type == "heading_open":
                level = int(tok.tag[1]) if tok.tag.startswith("h") and tok.tag[1:].isdigit() else None
                # Next token should be inline with content
                if i + 1 < len(tokens) and tokens[i + 1].type == "inline":
                    inline_token = tokens[i + 1]
                    heading_text = inline_token.content or ""
                    formatted_runs = _parse_inline_formatting(inline_token)
                    blocks.append(
                        MarkdownBlock(
                            type="heading",
                            text=heading_text,  # Keep for fallback
                            formatted_runs=formatted_runs if formatted_runs else None,
                            level=level,
                        )
                    )
                # Skip until heading_close
                while i < len(tokens) and tokens[i].type != "heading_close":
                    i += 1
                i += 1
                continue


            # Bullet lists (unordered) with nesting support
            if tok.type == "bullet_list_open":
                items: List[tuple[int, str]] = []
                
                def parse_bullet_list(start_idx: int, base_level: int) -> int:
                    """
                    Parse a bullet list starting at start_idx.
                    Returns index after the matching bullet_list_close.
                    base_level: nesting level (1 for top-level, 2 for nested, etc.)
                    """
                    idx = start_idx
                    current_level = base_level
                    
                    while idx < len(tokens):
                        if tokens[idx].type == "bullet_list_close":
                            return idx + 1
                        
                        if tokens[idx].type == "bullet_list_open":
                            # Nested list - recursively parse it
                            idx = parse_bullet_list(idx + 1, current_level + 1)
                            continue
                        
                        if tokens[idx].type == "list_item_open":
                            # Collect text from first paragraph of this list item
                            item_text_parts: List[str] = []
                            inline_token = None
                            j = idx + 1
                            found_para_content = False
                            
                            while j < len(tokens) and tokens[j].type != "list_item_close":
                                if tokens[j].type == "paragraph_open":
                                    found_para_content = False  # Reset for new paragraph
                                elif tokens[j].type == "inline":
                                    if tokens[j].content:
                                        item_text_parts.append(tokens[j].content)
                                        found_para_content = True
                                    if inline_token is None:
                                        inline_token = tokens[j]  # Store for formatting parsing
                                elif tokens[j].type == "paragraph_close":
                                    # Stop after first paragraph
                                    break
                                elif tokens[j].type == "bullet_list_open":
                                    # Nested list found - stop collecting text
                                    break
                                j += 1
                            
                            item_text = " ".join(p for p in item_text_parts if p)
                            if item_text.strip():
                                # Preserve actual logical level from markdown (don't clamp during parsing)
                                # Clamping will happen during rendering based on config
                                logical_level = current_level
                                # Parse formatting if available
                                formatted_runs = None
                                if inline_token:
                                    formatted_runs = _parse_inline_formatting(inline_token)
                                    if not formatted_runs:
                                        formatted_runs = None
                                items.append((logical_level, item_text, formatted_runs))
                            
                            # Skip to list_item_close, processing any nested lists
                            while idx < len(tokens) and tokens[idx].type != "list_item_close":
                                if tokens[idx].type == "bullet_list_open":
                                    idx = parse_bullet_list(idx + 1, current_level + 1)
                                    continue
                                idx += 1
                            idx += 1
                            continue
                        
                        idx += 1
                    
                    return idx
                
                # Parse the list starting from the next token
                i = parse_bullet_list(i + 1, 1)
                
                if items:
                    blocks.append(
                        MarkdownBlock(
                            type="bullet_list",
                            items=items,
                        )
                    )
                continue

            # Numbered lists (ordered) with nesting support
            if tok.type == "ordered_list_open":
                items: List[tuple[int, str, Optional[List[FormattedRun]]]] = []
                logged_nested_bullet = False
                
                def parse_ordered_list(start_idx: int, base_level: int) -> int:
                    """
                    Parse an ordered list starting at start_idx.
                    Returns index after the matching ordered_list_close.
                    base_level: nesting level (1 for top-level, 2 for nested, etc.)
                    """
                    nonlocal logged_nested_bullet
                    idx = start_idx
                    current_level = base_level
                    
                    while idx < len(tokens):
                        if tokens[idx].type == "ordered_list_close":
                            return idx + 1
                        
                        if tokens[idx].type == "ordered_list_open":
                            # Nested list - recursively parse it
                            idx = parse_ordered_list(idx + 1, current_level + 1)
                            continue
                        
                        if tokens[idx].type == "list_item_open":
                            # Collect text from first paragraph of this list item
                            item_text_parts: List[str] = []
                            inline_token = None
                            j = idx + 1
                            
                            while j < len(tokens) and tokens[j].type != "list_item_close":
                                if tokens[j].type == "paragraph_open":
                                    pass  # Reset for new paragraph
                                elif tokens[j].type == "inline":
                                    if tokens[j].content:
                                        item_text_parts.append(tokens[j].content)
                                    if inline_token is None:
                                        inline_token = tokens[j]  # Store for formatting parsing
                                elif tokens[j].type == "paragraph_close":
                                    # Stop after first paragraph
                                    break
                                elif tokens[j].type in ("ordered_list_open", "bullet_list_open"):
                                    # Nested list found - stop collecting text
                                    break
                                j += 1
                            
                            item_text = " ".join(p for p in item_text_parts if p)
                            if item_text.strip():
                                # Preserve actual logical level from markdown (don't clamp during parsing)
                                # Clamping will happen during rendering based on config
                                logical_level = current_level
                                # Parse formatting if available
                                formatted_runs = None
                                if inline_token:
                                    formatted_runs = _parse_inline_formatting(inline_token)
                                    if not formatted_runs:
                                        formatted_runs = None
                                items.append((logical_level, item_text, formatted_runs))
                            
                            # Skip to list_item_close, processing any nested lists
                            while idx < len(tokens) and tokens[idx].type != "list_item_close":
                                if tokens[idx].type in ("ordered_list_open", "bullet_list_open"):
                                    # Recursively parse nested list
                                    if tokens[idx].type == "ordered_list_open":
                                        idx = parse_ordered_list(idx + 1, current_level + 1)
                                    else:
                                        # Nested bullet list - skip for now (could handle if needed)
                                        if not logged_nested_bullet:
                                            logger.debug(
                                                "Nested bullet list inside ordered list skipped for determinism."
                                            )
                                            logged_nested_bullet = True
                                        nested_depth = 1
                                        k = idx + 1
                                        while k < len(tokens) and nested_depth > 0:
                                            if tokens[k].type == "bullet_list_open":
                                                nested_depth += 1
                                            elif tokens[k].type == "bullet_list_close":
                                                nested_depth -= 1
                                            k += 1
                                        idx = k
                                    continue
                                idx += 1
                            idx += 1
                            continue
                        
                        idx += 1
                    
                    return idx
                
                # Parse the list starting from the next token
                i = parse_ordered_list(i + 1, 1)
                
                if items:
                    blocks.append(
                        MarkdownBlock(
                            type="numbered_list",
                            items=items,
                        )
                    )
                continue

            # Tables
            if tok.type == "table_open":
                header: List[str] = []
                rows: List[List[str]] = []

                i += 1
                while i < len(tokens) and tokens[i].type != "table_close":
                    # Header row
                    if tokens[i].type == "thead_open":
                        i += 1
                        while i < len(tokens) and tokens[i].type != "thead_close":
                            if tokens[i].type == "tr_open":
                                i += 1
                                current_row: List[str] = []
                                while i < len(tokens) and tokens[i].type != "tr_close":
                                    if tokens[i].type == "th_open":
                                        # Expect inline next
                                        if i + 1 < len(tokens) and tokens[i + 1].type == "inline":
                                            cell_text = tokens[i + 1].content or ""
                                            current_row.append(cell_text)
                                        # Skip to th_close
                                        while i < len(tokens) and tokens[i].type != "th_close":
                                            i += 1
                                    i += 1
                                if current_row:
                                    header = current_row
                            i += 1
                        i += 1
                        continue

                    # Body rows
                    if tokens[i].type == "tbody_open":
                        i += 1
                        while i < len(tokens) and tokens[i].type != "tbody_close":
                            if tokens[i].type == "tr_open":
                                i += 1
                                current_row: List[str] = []
                                while i < len(tokens) and tokens[i].type != "tr_close":
                                    if tokens[i].type == "td_open":
                                        if i + 1 < len(tokens) and tokens[i + 1].type == "inline":
                                            cell_text = tokens[i + 1].content or ""
                                            current_row.append(cell_text)
                                        # Skip to td_close
                                        while i < len(tokens) and tokens[i].type != "td_close":
                                            i += 1
                                    i += 1
                                if current_row:
                                    rows.append(current_row)
                            i += 1
                        i += 1
                        continue

                    i += 1

                if header or rows:
                    blocks.append(
                        MarkdownBlock(
                            type="table",
                            header=header or [],
                            rows=rows or [],
                        )
                    )
                continue

            i += 1

        # If parsing produced no blocks but text is non-empty, fall back to paragraphs
        if not blocks:
            logger.debug("Markdown parsed but yielded no blocks; falling back to paragraphs.")
            raise ValueError("No markdown blocks produced")

        return blocks

    except Exception as e:
        # Fail-safe: treat content as plain text paragraphs separated by blank lines
        logger.warning(f"Markdown parsing failed, falling back to plain paragraphs: {e}")
        paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
        return [MarkdownBlock(type="paragraph", text=p) for p in paragraphs]


def _iter_header_footer_paragraphs(doc: DocumentType) -> Iterable[Tuple[Paragraph, Region]]:
    """
    Yield all paragraphs in headers and footers, including table cells.
    
    Returns tuples of (paragraph, region) where region is HEADER or FOOTER.
    """
    for section in doc.sections:
        # Header paragraphs
        for p in section.header.paragraphs:
            yield (p, Region.HEADER)
        
        # Header table cells
        for table in section.header.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield (p, Region.HEADER)
        
        # Footer paragraphs
        for p in section.footer.paragraphs:
            yield (p, Region.FOOTER)
        
        # Footer table cells
        for table in section.footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield (p, Region.FOOTER)


def _iter_textbox_runs(doc: DocumentType, region: Optional[Region] = None) -> Iterable[Tuple[List[Any], Region, Container]]:
    """
    Iterate through textbox runs using XML/XPath.
    
    Yields (runs_list, region, TEXTBOX) tuples where runs_list contains
    XML run elements (w:r) from a single textbox.
    
    Args:
        doc: Document to traverse
        region: Region to search (BODY, HEADER, or FOOTER)
    
    Returns:
        Iterator of (runs_list, region, Container.TEXTBOX) tuples
    """
    from lxml import etree
    
    # Namespace mapping for XPath
    ns = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    }

    # When region is None, iterate all regions in a fixed order.
    if region is None:
        for reg in (Region.BODY, Region.HEADER, Region.FOOTER):
            for runs, textbox_region, container in _iter_textbox_runs(doc, reg):
                yield (runs, textbox_region, container)
        return
    
    # Determine which XML element(s) to search based on region
    if region == Region.BODY:
        # Search the main document body
        search_element = doc._body._element
        # Use lxml etree xpath directly (supports namespaces)
        for txbx_content in etree._Element.xpath(search_element, './/wps:txbx//w:txbxContent', namespaces=ns):
            # Collect all runs within this textbox
            runs = etree._Element.xpath(txbx_content, './/w:r', namespaces=ns)
            if runs:
                yield (runs, region, Container.TEXTBOX)
    elif region == Region.HEADER:
        # Search all header sections
        for section in doc.sections:
            header = section.header
            search_element = header._element
            # Use lxml etree xpath directly.
            # NOTE: `.//w:txbxContent` is a superset that covers both WPS and classic/VML textboxes.
            for txbx_content in etree._Element.xpath(search_element, './/w:txbxContent', namespaces=ns):
                # Collect all runs within this textbox
                runs = etree._Element.xpath(txbx_content, './/w:r', namespaces=ns)
                if runs:
                    yield (runs, region, Container.TEXTBOX)
    elif region == Region.FOOTER:
        # Search all footer sections
        for section in doc.sections:
            footer = section.footer
            search_element = footer._element
            # Use lxml etree xpath directly.
            # NOTE: `.//w:txbxContent` is a superset that covers both WPS and classic/VML textboxes.
            for txbx_content in etree._Element.xpath(search_element, './/w:txbxContent', namespaces=ns):
                # Collect all runs within this textbox
                runs = etree._Element.xpath(txbx_content, './/w:r', namespaces=ns)
                if runs:
                    yield (runs, region, Container.TEXTBOX)


def _iter_all_locations(doc: DocumentType) -> Iterable[Tuple[Union[Paragraph, List[Any]], Region, Container]]:
    """
    Yield all content locations in the document with location metadata.
    
    Returns tuples of (paragraph_or_runs, region, container):
    - paragraph_or_runs: Paragraph object for FLOW/TABLE, or list of XML run elements for TEXTBOX
    - region: BODY, HEADER, or FOOTER
    - container: FLOW (normal paragraphs), TABLE (table cells), or TEXTBOX (w:txbxContent)
    """
    # Body paragraphs (FLOW)
    for p in doc.paragraphs:
        logger.debug("Traversing BODY FLOW paragraph")
        yield (p, Region.BODY, Container.FLOW)
    
    # Body table cells (TABLE)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    logger.debug("Traversing BODY TABLE cell paragraph")
                    yield (p, Region.BODY, Container.TABLE)
    
    # Header and footer paragraphs and table cells
    for p, region in _iter_header_footer_paragraphs(doc):
        # Determine if this is FLOW or TABLE
        # Check if paragraph's parent is a table cell
        parent = p._p.getparent()
        is_table = parent is not None and parent.tag.endswith('}tc')  # w:tc element
        container = Container.TABLE if is_table else Container.FLOW
        logger.debug(f"Traversing {region.value} {container.value} paragraph")
        yield (p, region, container)
    
    # Textboxes in all regions
    for region in [Region.BODY, Region.HEADER, Region.FOOTER]:
        for runs, textbox_region, container in _iter_textbox_runs(doc, region):
            logger.debug(f"Traversing {textbox_region.value} TEXTBOX with {len(runs)} runs")
            yield (runs, textbox_region, container)


def _is_block_expansion_allowed(region: Region, container: Container) -> bool:
    """
    Check if block expansion is allowed for the given location.
    
    Block expansion is only allowed in BODY FLOW and BODY TABLE.
    """
    return region == Region.BODY and container in [Container.FLOW, Container.TABLE]


def _get_text_from_xml_run(run_elem) -> str:
    """
    Extract text from an XML run element (w:r).
    
    Returns the concatenated text from all w:t elements within the run.
    """
    from lxml import etree
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    text_elems = etree._Element.xpath(run_elem, './/w:t', namespaces=ns)
    return "".join(elem.text or "" for elem in text_elems)


def _set_text_in_xml_run(run_elem, text: str) -> None:
    """
    Set text in an XML run element (w:r), preserving structure.
    
    If the run has w:t elements, updates the first one and removes others.
    If no w:t elements exist, creates one.
    """
    from lxml import etree
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    text_elems = etree._Element.xpath(run_elem, './/w:t', namespaces=ns)
    
    if text_elems:
        # Update first text element; do not remove nodes (stability invariant)
        text_elems[0].text = text
        for elem in text_elems[1:]:
            elem.text = ""
    else:
        # No existing w:t; do not create nodes (stability invariant)
        return


def _safe_replace_scalar_across_textbox_runs(
    run_elems: List[Any],
    placeholder: str,
    replacement: str,
    region: Region,
) -> None:
    """
    Safely replace a scalar placeholder inside textbox runs, supporting placeholders
    that span multiple XML runs (w:r).

    Rules:
    - Replacement can span multiple runs
    - No run insertion or deletion
    - No formatting changes (w:rPr preserved)
    - Only text content is modified
    - If placeholder found more than once or span cannot be mapped cleanly,
      leave content unchanged and log a warning.

    Args:
        run_elems: List of XML run elements (w:r)
        placeholder: Placeholder text to replace (e.g., "{{key}}")
        replacement: Replacement text
        region: Document region (BODY, HEADER, FOOTER) for logging
    """
    if not run_elems:
        return

    # Read-only scan: Extract text from each run and build concatenated string
    run_texts = []
    for run in run_elems:
        text = _get_text_from_xml_run(run)
        run_texts.append(text)
    
    full_text = "".join(run_texts)
    
    # Detection: If placeholder not found, return
    if placeholder not in full_text:
        return
    
    # Detection: If placeholder found more than once, log debug and return
    # (This should not happen as we check in replace_scalar_placeholders, but defensive check)
    placeholder_count = full_text.count(placeholder)
    if placeholder_count != 1:
        logger.debug(
            "Skipping textbox placeholder %s in %s TEXTBOX (found %d times, ambiguous). Leaving unchanged.",
            placeholder,
            region.value,
            placeholder_count,
        )
        return
    
    # Span mapping: Determine start run index + offset and end run index + offset
    placeholder_start = full_text.find(placeholder)
    placeholder_end = placeholder_start + len(placeholder)
    
    # Map placeholder position to runs
    current_pos = 0
    start_run_idx = None
    start_offset = None
    end_run_idx = None
    end_offset = None
    
    for i, run_text in enumerate(run_texts):
        run_start = current_pos
        run_end = current_pos + len(run_text)
        
        # Check if placeholder starts in this run
        if start_run_idx is None and run_start <= placeholder_start < run_end:
            start_run_idx = i
            start_offset = placeholder_start - run_start
        
        # Check if placeholder ends in this run
        if run_start < placeholder_end <= run_end:
            end_run_idx = i
            end_offset = placeholder_end - run_start
            break
        
        current_pos = run_end
    
    # Abort on uncertainty: If span cannot be mapped cleanly, log and return unchanged
    if start_run_idx is None or end_run_idx is None:
        logger.warning(
            "Skipping textbox placeholder %s in %s TEXTBOX (cannot map span cleanly). Leaving unchanged.",
            placeholder,
            region.value,
        )
        return
    
    # Safe rewrite: Modify text ONLY
    # Start run: prefix + replacement
    # Middle runs: set text to empty string
    # End run: suffix
    
    # Update start run
    start_run = run_elems[start_run_idx]
    start_run_text = run_texts[start_run_idx]
    start_prefix = start_run_text[:start_offset] if start_offset > 0 else ""
    start_new_text = start_prefix + replacement
    _set_text_in_xml_run(start_run, start_new_text)
    
    # Update middle runs (if any)
    for i in range(start_run_idx + 1, end_run_idx):
        middle_run = run_elems[i]
        _set_text_in_xml_run(middle_run, "")
    
    # Update end run (if different from start run)
    if end_run_idx > start_run_idx:
        end_run = run_elems[end_run_idx]
        end_run_text = run_texts[end_run_idx]
        end_suffix = end_run_text[end_offset:] if end_offset < len(end_run_text) else ""
        _set_text_in_xml_run(end_run, end_suffix)
    
    logger.debug(
        "Successfully replaced scalar placeholder %s across %d run(s) in %s TEXTBOX",
        placeholder,
        end_run_idx - start_run_idx + 1,
        region.value,
    )


def _replace_text_in_runs(
    runs: List[Any],
    placeholder: str,
    replacement: str,
) -> None:
    """
    Replace all occurrences of `placeholder` with `replacement` across a sequence
    of runs, handling cases where the placeholder spans multiple runs.

    This function preserves run count and styling, modifying only text content.
    """
    if not runs:
        return

    # Build concatenated text from existing XML text nodes only.
    # NOTE: We intentionally do NOT use `run.text = ...` anywhere in this function,
    # because python-docx may create/remove `w:t` nodes, which is forbidden.
    # We only mutate existing `w:t.text` via `_set_text_in_xml_run`.
    full_text = "".join(_get_text_from_xml_run(run._r) for run in runs)
    if not full_text or placeholder not in full_text:
        return

    # Idempotent + deterministic safety rule:
    # If the placeholder appears more than once, skip entirely (no partial edits).
    count = full_text.count(placeholder)
    if count != 1:
        logger.debug(
            "Skipping scalar placeholder %s across runs (found %d times, ambiguous). Leaving unchanged.",
            placeholder,
            count,
        )
        return

    idx = full_text.find(placeholder)
    end = idx + len(placeholder)

    new_texts: List[str] = []
    current_start = 0
    inserted = False

    # Compute the rewritten text per run without mutating anything yet.
    for run in runs:
        original = _get_text_from_xml_run(run._r)
        run_len = len(original)
        run_start = current_start
        run_end = current_start + run_len

        # No overlap with placeholder span
        if run_end <= idx or run_start >= end:
            new_texts.append(original)
        else:
            before_len = max(0, idx - run_start)
            after_len = max(0, run_end - end)

            before = original[:before_len] if before_len > 0 else ""
            after = original[run_len - after_len :] if after_len > 0 else ""

            middle = ""
            if not inserted:
                middle = replacement
                inserted = True

            new_texts.append(before + middle + after)

        current_start += run_len

    # Transactional guard: we only mutate runs that already have at least one `w:t`.
    # If any run would need changing but has no `w:t`, we skip entirely.
    from lxml import etree
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    for run, new_text in zip(runs, new_texts):
        current_text = _get_text_from_xml_run(run._r)
        if new_text == current_text:
            continue
        text_elems = etree._Element.xpath(run._r, ".//w:t", namespaces=ns)
        if not text_elems:
            logger.debug(
                "Skipping scalar placeholder %s across runs (run has no w:t; no node creation allowed). Leaving unchanged.",
                placeholder,
            )
            return

    # Commit: text-only mutation of existing XML nodes.
    for run, new_text in zip(runs, new_texts):
        _set_text_in_xml_run(run._r, new_text)


def _replace_header_footer_textboxes(doc: DocumentType, concrete: Dict[str, str]) -> None:
    """
    Phase 1 (MUST RUN FIRST): Replace scalar placeholders inside HEADER/FOOTER textboxes only.

    Transactional rules:
    - Read-only preflight (string only)
    - Replace only if:
      - placeholder appears exactly once in the textbox
      - replacement is non-empty
      - placeholder may span multiple XML runs; replacement is span-safe and deterministic
    - Commit mutates only w:t.text values (no node creation/deletion, no relocation)
    - If unsafe, do nothing
    """
    # Use unified textbox discovery; HEADER/FOOTER textboxes are filtered here.
    for run_elems, region, container in _iter_textbox_runs(doc):
        if region not in (Region.HEADER, Region.FOOTER):
            continue
        if container is not Container.TEXTBOX:
            continue

        # Build concatenated text from XML runs for detection
        full_text = "".join(_get_text_from_xml_run(run) for run in run_elems)
        if not full_text:
            continue

        # Only replace tokens that are in the authoritative mapping (scalar placeholders)
        for token, replacement in concrete.items():
            if token not in full_text:
                continue

            # Never replace empty or null values
            if not replacement or not replacement.strip():
                continue

            # Must appear exactly once
            if full_text.count(token) != 1:
                continue

            # SAFE cross-run replacement (no structural mutation)
            _safe_replace_scalar_across_textbox_runs(
                run_elems,
                token,
                replacement,
                region,
            )


def _replace_header_footer_flow_and_tables(doc: DocumentType, concrete: Dict[str, str]) -> None:
    """
    Phase 3: Replace scalar placeholders in HEADER/FOOTER normal text (paragraphs + table cells),
    explicitly skipping anything inside w:txbxContent.
    """
    from lxml import etree
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    for paragraph, region in _iter_header_footer_paragraphs(doc):
        # Explicitly skip any paragraph that is inside a textbox content subtree
        try:
            if etree._Element.xpath(paragraph._p, "ancestor::w:txbxContent", namespaces=ns):
                continue
        except Exception:
            # If we can't determine safely, skip.
            continue

        if not paragraph.runs:
            continue

        para_text = "".join(run.text or "" for run in paragraph.runs)
        if not para_text:
            continue

        for token, replacement in concrete.items():
            if not replacement or not replacement.strip():
                continue
            if token in para_text:
                _replace_text_in_runs(paragraph.runs, token, replacement)
                para_text = "".join(run.text or "" for run in paragraph.runs)


def replace_scalar_placeholders(doc: DocumentType, mapping: Dict[str, Optional[str]]) -> None:
    """
    Replace scalar placeholders (e.g. {{document_id}}) in all locations:
    - Body paragraphs and table cells
    - Header paragraphs and table cells
    - Footer paragraphs and table cells
    - Textboxes (in body, headers, and footers)

    This function only replaces inline text and never inserts or removes paragraphs.
    
    ⚠️  PLACEHOLDER PROVENANCE: Only placeholders present in the `mapping` parameter
    will be replaced. Any placeholder found in the document but not in `mapping` will
    be ignored (logged at debug level).
    """
    # Build concrete placeholder -> replacement mapping
    # This enforces authoritative placeholder registry - only keys from mapping are replaced
    concrete: Dict[str, Optional[str]] = {}
    for key, value in mapping.items():
        token = f"{{{{{key}}}}}"  # e.g. document_id -> {{document_id}}
        concrete[token] = value

    if not concrete:
        return

    # Build set of allowed placeholders for provenance checking
    allowed_tokens = set(concrete.keys())

    # Phase 1 — HEADER & FOOTER TEXTBOXES (XML-only, FIRST)
    _replace_header_footer_textboxes(doc, {k: (v or "") for k, v in concrete.items()})

    # Phase 2 — BODY SCALAR REPLACEMENT (Unchanged)
    for location_data, region, container in _iter_all_locations(doc):
        if region != Region.BODY:
            continue

        if container == Container.TEXTBOX:
            run_elems = location_data  # list of XML w:r elements
            if not run_elems:
                continue

            full_text = "".join(_get_text_from_xml_run(run) for run in run_elems)
            if not full_text:
                continue

            for token, replacement in concrete.items():
                if token not in full_text:
                    continue
                if not replacement or not replacement.strip():
                    logger.debug(
                        "Skipping %s TEXTBOX replacement for %s (empty replacement)",
                        region.value,
                        token,
                    )
                    continue
                if full_text.count(token) != 1:
                    logger.debug(
                        "Skipping %s TEXTBOX replacement for %s (ambiguous occurrence)",
                        region.value,
                        token,
                    )
                    continue

                _safe_replace_scalar_across_textbox_runs(run_elems, token, replacement, region)

        else:
            paragraph = location_data  # Paragraph in BODY FLOW/TABLE
            if not paragraph.runs:
                continue
            para_text = "".join(run.text or "" for run in paragraph.runs)
            if not para_text:
                continue

            for token, replacement in concrete.items():
                if not replacement or not replacement.strip():
                    continue
                if token in para_text:
                    logger.debug(f"Replacing scalar placeholder {token} in {region.value} {container.value}")
                    _replace_text_in_runs(paragraph.runs, token, replacement)
                    para_text = "".join(run.text or "" for run in paragraph.runs)

    # Phase 3 — HEADER & FOOTER NON-TEXTBOX TEXT (skip w:txbxContent)
    _replace_header_footer_flow_and_tables(doc, concrete)


def _is_paragraph_in_table_cell(paragraph: Paragraph) -> bool:
    """Check if a paragraph is inside a table cell."""
    parent = paragraph._p.getparent()
    if parent is None:
        return False
    return parent.tag.endswith('}tc')  # w:tc element


def _get_cell_for_paragraph(paragraph: Paragraph):
    """Get the table cell that contains this paragraph, or None."""
    parent = paragraph._parent
    if hasattr(parent, "_tc"):
        return parent
    return None


def _set_paragraph_formatted_text(paragraph: Paragraph, formatted_runs: Optional[List[FormattedRun]], fallback_text: str = "") -> None:
    """
    Set paragraph text using formatted runs if available, otherwise use fallback text.
    Clears existing runs first.
    """
    # Clear existing runs
    paragraph.clear()
    
    if formatted_runs:
        # Add formatted runs
        for run_data in formatted_runs:
            run = paragraph.add_run(run_data.text)
            run.bold = run_data.bold
            run.italic = run_data.italic
    elif fallback_text:
        paragraph.text = fallback_text


def _heading_style_for_level(level: Optional[int]) -> str:
    if level is None or level < 1:
        return "Heading 1"
    if level > 9:
        level = 9
    return f"Heading {level}"


def _try_apply_word_table_style(tbl_elem, doc_body):
    """
    Attempt to apply a Word table style to the table element.
    
    Args:
        tbl_elem: XML element (w:tbl) to format
        doc_body: Document body element for style application via Table object
    
    Returns:
        Table object if style was successfully applied, None otherwise
    """
    if doc_body is None:
        return None
    
    try:
        from docx.table import Table
        table = Table(tbl_elem, doc_body)
        
        # Try styles in preferred order
        for style_name in ['Light Grid', 'Medium Grid 1', 'Table Grid']:
            try:
                table.style = style_name
                return table
            except Exception:
                continue
        
        return None
    except Exception:
        return None


def _apply_header_row_emphasis(table_obj):
    """
    Make header row (first row) text bold.
    
    Args:
        table_obj: Table object from python-docx
    """
    if len(table_obj.rows) > 0:
        header_row = table_obj.rows[0]
        for cell in header_row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.bold = True


def _apply_xml_table_borders_fallback(tbl_elem):
    """
    Apply explicit XML borders as fallback when Word style cannot be applied.
    
    Args:
        tbl_elem: XML element (w:tbl) to format
    """
    from docx.oxml.ns import qn
    
    # Find or create w:tblPr element
    tbl_pr = None
    for child in tbl_elem:
        if child.tag == qn('w:tblPr'):
            tbl_pr = child
            break
    
    if tbl_pr is None:
        # Create tblPr if it doesn't exist
        tbl_pr = tbl_elem.makeelement(qn('w:tblPr'))
        # Insert at the beginning
        tbl_elem.insert(0, tbl_pr)
    
    # Check if w:tblBorders already exists
    tbl_borders = None
    for child in tbl_pr:
        if child.tag == qn('w:tblBorders'):
            tbl_borders = child
            break
    
    # If borders already exist, skip (idempotent behavior)
    if tbl_borders is not None and len(tbl_borders) > 0:
        return
    
    if tbl_borders is None:
        # Create tblBorders element
        tbl_borders = tbl_pr.makeelement(qn('w:tblBorders'))
        tbl_pr.append(tbl_borders)
    
    # Define border properties: single line, size 4 (0.5pt), auto color
    border_attrs = {
        qn('w:val'): 'single',
        qn('w:sz'): '4',
        qn('w:space'): '0',
        qn('w:color'): 'auto',
    }
    
    # Apply borders to all sides
    border_names = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
    for border_name in border_names:
        border_elem = tbl_borders.makeelement(qn(f'w:{border_name}'))
        for attr_name, attr_value in border_attrs.items():
            border_elem.set(attr_name, attr_value)
        tbl_borders.append(border_elem)


def _apply_default_table_formatting(tbl_elem, doc_body=None):
    """
    Apply default visible grid formatting to a Word table XML element.
    
    First attempts to apply built-in Word styles in order: "Light Grid", "Medium Grid 1", "Table Grid".
    If style assignment succeeds, also makes header row text bold.
    If style assignment fails or doc_body is None, falls back to explicit XML borders.
    
    Args:
        tbl_elem: XML element (w:tbl) to format
        doc_body: Optional document body element for style application via Table object
    """
    table_obj = _try_apply_word_table_style(tbl_elem, doc_body)
    if table_obj is not None:
        _apply_header_row_emphasis(table_obj)
        return
    
    _apply_xml_table_borders_fallback(tbl_elem)


def _attempt_render_nested_word_table(block, *, parent_cell_elem, insertion_index, render_context) -> bool:
    """
    Attempt to render a markdown table block as a real nested Word table
    inside a table cell, with strict validation and atomic commit.

    Returns True only when the nested table is built, validated, and inserted.
    Returns False (with INFO log) on any failure; must not raise.
    """
    # INVARIANT:
    # Nested table rendering is strictly atomic:
    # - Build in scratch document
    # - Validate fully
    # - Clone XML
    # - Insert exactly once
    # Any failure MUST result in zero document mutation.
    
    from copy import deepcopy
    from docx import Document
    from docx.oxml.ns import qn
    from lxml import etree

    try:
        header = block.header or []
        rows = block.rows or []

        # Determine number of columns
        num_cols = len(header) if header else (len(rows[0]) if rows else 0)
        if num_cols <= 0:
            logger.info("Nested table could not be rendered safely; falling back to text grid.")
            return False

        # Build the table in isolation using a scratch document
        scratch_doc = Document()
        scratch_table = scratch_doc.add_table(rows=0, cols=num_cols)

        # Populate header row (if any)
        if header:
            hdr_row = scratch_table.add_row()
            for idx, cell_text in enumerate(header[:num_cols]):
                hdr_row.cells[idx].text = str(cell_text) if cell_text else ""

        # Populate data rows
        for row in rows:
            data_row = scratch_table.add_row()
            for idx, cell_text in enumerate(row[:num_cols]):
                data_row.cells[idx].text = str(cell_text) if cell_text else ""

        # Apply formatting BEFORE cloning (using scratch document context for style resolution)
        _apply_default_table_formatting(scratch_table._tbl, scratch_doc._body)

        # Extract and clone the already-formatted w:tbl element
        tbl_elem = scratch_table._tbl
        cloned_tbl = deepcopy(tbl_elem)

        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        # Validate: root element must be a table
        if not cloned_tbl.tag.endswith("}tbl"):
            logger.info("Nested table could not be rendered safely; falling back to text grid.")
            return False

        # Disallow section properties or numbering-related XML inside the nested table subtree
        disallowed_xpaths = [
            ".//w:sectPr",
            ".//w:numPr",
            ".//w:abstractNum",
        ]
        for xp in disallowed_xpaths:
            if etree._Element.xpath(cloned_tbl, xp, namespaces=ns):
                logger.info("Nested table could not be rendered safely; falling back to text grid.")
                return False

        # Validate parent container is exactly w:tc
        if parent_cell_elem is None or parent_cell_elem.tag != qn("w:tc"):
            logger.error(
                "Nested table parent is not w:tc; tag=%s",
                parent_cell_elem.tag if parent_cell_elem is not None else None,
            )
            logger.info("Nested table could not be rendered safely; falling back to text grid.")
            return False

        # Atomic commit: ensure there are paragraph siblings, then insert at position.
        # The original placeholder paragraph has already been removed earlier.
        try:
            # Ensure at least one <w:p> exists to satisfy Word rendering expectations.
            has_any_para = any(child.tag == qn("w:p") for child in parent_cell_elem)
            if not has_any_para:
                from docx.oxml import parse_xml
                from docx.oxml.ns import nsdecls

                p_xml = parse_xml(f'<w:p {nsdecls("w")}/>')
                parent_cell_elem.insert(insertion_index, p_xml)
                insertion_index += 1

            # Insert the validated table
            parent_cell_elem.insert(insertion_index, cloned_tbl)

            # Ensure there is also a paragraph AFTER the table so that the
            # cell structure is: <w:p/> <w:tbl/> <w:p/>.
            children = list(parent_cell_elem)
            tbl_index = children.index(cloned_tbl)
            if tbl_index == len(children) - 1:
                from docx.oxml import parse_xml
                from docx.oxml.ns import nsdecls

                p_xml_after = parse_xml(f'<w:p {nsdecls("w")}/>')
                parent_cell_elem.insert(tbl_index + 1, p_xml_after)
        except Exception:
            logger.info("Nested table could not be rendered safely; falling back to text grid.")
            return False

        logger.info(
            "Nested Word table committed into table cell at index %s",
            insertion_index,
        )
        return True

    except Exception:
        logger.info(
            "Nested table could not be rendered safely; falling back to text grid.",
            exc_info=True,
        )
        return False


def replace_block_placeholder_with_content(
    doc: DocumentType,
    paragraph: Paragraph,
    placeholder_token: str,
    blocks: Optional[List[MarkdownBlock]],
    config: "ListRenderConfig",
) -> None:
    """
    Replace a block placeholder (e.g. {{summary}}) in the given paragraph with
    structured content (paragraphs, headings, bullet lists, tables).

    The placeholder text itself is removed from the paragraph before inserting
    any new content.
    """
    # INVARIANT:
    # Paragraph insertion is performed inline (no shared helper) to preserve
    # exact XML insertion order and avoid accidental cross-container mutation.
    # This duplication is intentional for determinism and safety.
    
    # Defensive guard: ensure config is never None
    if config is None:
        raise ValueError("config must not be None")
    
    # Determine container: BODY FLOW or TABLE CELL
    # Get cell reference and actual XML parent BEFORE removing paragraph
    actual_parent_elem = paragraph._p.getparent()
    cell = _get_cell_for_paragraph(paragraph)
    is_in_table_cell = cell is not None
    cell_elem = actual_parent_elem if is_in_table_cell and actual_parent_elem is not None else None  # type: ignore
    
    # Enforce "block placeholder must occupy entire paragraph" rule
    # Reconstruct full paragraph text to check if placeholder is the only content
    para_text = "".join(run.text or "" for run in paragraph.runs)
    if para_text.strip() != placeholder_token:
        logger.warning(
            f"Block placeholder {placeholder_token} is not the only content in its paragraph. "
            f"Paragraph contains: {para_text[:100]}. "
            "Block placeholders must occupy the entire paragraph. Skipping replacement."
        )
        return
    
    # Capture insertion point BEFORE removing paragraph
    # This ensures new content is inserted at the exact position of the placeholder
    insertion_parent = None
    insertion_index = None
    cell_insertion_index = None
    if is_in_table_cell and cell_elem is not None:
        cell_insertion_index = list(cell_elem).index(paragraph._p)
    else:
        if actual_parent_elem is not None:
            insertion_parent = actual_parent_elem
            insertion_index = actual_parent_elem.index(paragraph._p)

    # Remove the placeholder paragraph from its parent container.
    # For BODY FLOW we remove it immediately; for TABLE CELL we may delay removal
    # to use it as an anchor for nested table insertion.
    placeholder_removed = False
    if not is_in_table_cell:
        paragraph._p.getparent().remove(paragraph._p)
        placeholder_removed = True
    
    # If no content blocks, we may still need to ensure the cell is not structurally empty
    if not blocks:
        if is_in_table_cell and cell_elem is not None:
            from lxml import etree
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            existing_paras = etree._Element.xpath(cell_elem, ".//w:p", namespaces=ns)
            if not existing_paras:
                from docx.oxml import parse_xml
                from docx.oxml.ns import nsdecls
                p_xml = parse_xml(f'<w:p {nsdecls("w")}/>')
                insert_at = cell_insertion_index if cell_insertion_index is not None else 0
                cell_elem.insert(insert_at, p_xml)
        return

    # Insert all new content at the correct position (BODY FLOW or TABLE CELL)
    for block in blocks:
        # For table-cell blocks other than tables, ensure the placeholder paragraph has been removed.
        if is_in_table_cell and block.type != "table" and not placeholder_removed:
            parent_for_para = paragraph._p.getparent()
            if parent_for_para is not None:
                parent_for_para.remove(paragraph._p)
            placeholder_removed = True

        if block.type == "heading":
            text = block.text or ""
            style = _heading_style_for_level(block.level)

            # Create paragraph directly in container
            if is_in_table_cell and cell_elem is not None:
                from docx.oxml import parse_xml
                from docx.oxml.ns import nsdecls
                p_xml = parse_xml(f'<w:p {nsdecls("w")}/>')
                insert_at = cell_insertion_index if cell_insertion_index is not None else len(cell_elem)
                cell_elem.insert(insert_at, p_xml)
                new_p = Paragraph(p_xml, cell)  # type: ignore[arg-type]
                cell_insertion_index = insert_at + 1
            else:
                # BODY FLOW: insert at captured position
                from docx.oxml import parse_xml
                from docx.oxml.ns import nsdecls
                p_xml = parse_xml(f'<w:p {nsdecls("w")}/>')
                insertion_parent.insert(insertion_index, p_xml)
                new_p = Paragraph(p_xml, doc._body)
                insertion_index += 1
            
            new_p.style = style
            _set_paragraph_formatted_text(new_p, block.formatted_runs, text)

        elif block.type == "paragraph":
            text = block.text or ""
            
            # Create paragraph directly in container
            if is_in_table_cell and cell_elem is not None:
                from docx.oxml import parse_xml
                from docx.oxml.ns import nsdecls
                p_xml = parse_xml(f'<w:p {nsdecls("w")}/>')
                insert_at = cell_insertion_index if cell_insertion_index is not None else len(cell_elem)
                cell_elem.insert(insert_at, p_xml)
                new_p = Paragraph(p_xml, cell)  # type: ignore[arg-type]
                cell_insertion_index = insert_at + 1
            else:
                # BODY FLOW: insert at captured position
                from docx.oxml import parse_xml
                from docx.oxml.ns import nsdecls
                p_xml = parse_xml(f'<w:p {nsdecls("w")}/>')
                insertion_parent.insert(insertion_index, p_xml)
                new_p = Paragraph(p_xml, doc._body)
                insertion_index += 1
            
            # Keep existing style (body style from template) - no style change for paragraphs
            _set_paragraph_formatted_text(new_p, block.formatted_runs, text)

        elif block.type == "bullet_list":
            # Manual bullet rendering - no Word numbering XML
            max_depth = config.max_visual_depth if config.max_visual_depth is not None else 3
            indent_per_level = config.indent_inches_per_level
            hanging_indent = -config.hanging_indent_inches
            
            # Guard against empty bullet_glyphs
            bullet_glyphs = config.bullet_glyphs if config.bullet_glyphs else ("•",)
            if not bullet_glyphs:
                bullet_glyphs = ("•",)

            raw_items = block.items or []
            clamped_logged = False  # Track if we've logged clamping for this block
            for idx, item in enumerate(raw_items):
                # Support legacy (str), (level, text), and new (level, text, formatted_runs) formats
                if isinstance(item, tuple):
                    if len(item) == 3:
                        logical_level, item_text, formatted_runs = item
                    elif len(item) == 2:
                        logical_level, item_text = item
                        formatted_runs = None
                    else:
                        logical_level, item_text = 1, item[0] if item else ""
                        formatted_runs = None
                else:
                    logical_level, item_text = 1, item
                    formatted_runs = None

                # Normalize logical level (actual nesting from markdown)
                if logical_level < 1:
                    logical_level = 1
                
                # Compute visual level (clamped by max_visual_depth) once per item
                if logical_level <= max_depth:
                    visual_level = logical_level
                    bullet_glyph = bullet_glyphs[min(logical_level - 1, len(bullet_glyphs) - 1)]
                else:
                    # Handle depth beyond max_visual_depth based on strategy
                    if config.deep_bullet_strategy == "clamp_last":
                        visual_level = max_depth
                        bullet_glyph = bullet_glyphs[min(max_depth - 1, len(bullet_glyphs) - 1)]
                        if not clamped_logged:
                            logger.debug("List nesting level %d clamped to %d", logical_level, max_depth)
                            clamped_logged = True
                    elif config.deep_bullet_strategy == "cycle":
                        # Cycle through available glyphs
                        visual_level = max_depth
                        glyph_index = ((logical_level - 1) % len(bullet_glyphs))
                        bullet_glyph = bullet_glyphs[glyph_index]
                    elif config.deep_bullet_strategy == "textual":
                        # Use textual indicator for deep nesting
                        visual_level = max_depth
                        bullet_glyph = f"[{logical_level}]"
                    else:
                        # Default to clamp_last
                        visual_level = max_depth
                        bullet_glyph = bullet_glyphs[min(max_depth - 1, len(bullet_glyphs) - 1)]
                        if not clamped_logged:
                            logger.debug("List nesting level %d clamped to %d", logical_level, max_depth)
                            clamped_logged = True

                # Create paragraph directly in container
                if is_in_table_cell and cell_elem is not None:
                    from docx.oxml import parse_xml
                    from docx.oxml.ns import nsdecls
                    p_xml = parse_xml(f'<w:p {nsdecls("w")}/>')
                    insert_at = cell_insertion_index if cell_insertion_index is not None else len(cell_elem)
                    cell_elem.insert(insert_at, p_xml)
                    new_p = Paragraph(p_xml, cell)  # type: ignore[arg-type]
                    cell_insertion_index = insert_at + 1
                else:
                    # BODY FLOW: insert at captured position
                    from docx.oxml import parse_xml
                    from docx.oxml.ns import nsdecls
                    p_xml = parse_xml(f'<w:p {nsdecls("w")}/>')
                    insertion_parent.insert(insertion_index, p_xml)
                    new_p = Paragraph(p_xml, doc._body)
                    insertion_index += 1

                # Set indentation using visual_level (consistent with glyph selection)
                new_p.paragraph_format.left_indent = Inches(indent_per_level * visual_level)
                new_p.paragraph_format.first_line_indent = Inches(hanging_indent)

                # Add bullet prefix as formatting-isolated run, then item text
                bullet_run = new_p.add_run(f"{bullet_glyph}  ")
                bullet_run.bold = False
                bullet_run.italic = False
                
                if formatted_runs:
                    for run_data in formatted_runs:
                        run = new_p.add_run(run_data.text)
                        run.bold = run_data.bold
                        run.italic = run_data.italic
                else:
                    new_p.add_run(item_text)

        elif block.type == "numbered_list":
            # Manual numbered list rendering - no Word numbering XML
            max_depth = config.max_visual_depth if config.max_visual_depth is not None else 3
            indent_per_level = config.indent_inches_per_level
            hanging_indent = -config.hanging_indent_inches
            
            # Track numbering per list block (resets for each block)
            # Use dictionary-based counters for dynamic depth (supports any nesting level)
            number_counters: Dict[int, int] = {}
            last_visual_level = 0

            raw_items = block.items or []
            clamped_logged = False  # Track if we've logged clamping for this block
            for idx, item in enumerate(raw_items):
                # Support legacy (str), (level, text), and new (level, text, formatted_runs) formats
                if isinstance(item, tuple):
                    if len(item) == 3:
                        logical_level, item_text, formatted_runs = item
                    elif len(item) == 2:
                        logical_level, item_text = item
                        formatted_runs = None
                    else:
                        logical_level, item_text = 1, item[0] if item else ""
                        formatted_runs = None
                else:
                    logical_level, item_text = 1, item
                    formatted_runs = None

                # Normalize logical level (actual nesting from markdown)
                if logical_level < 1:
                    logical_level = 1
                
                # Compute visual level (clamped by max_visual_depth) for rendering
                visual_level = min(logical_level, max_depth)
                if logical_level > max_depth and not clamped_logged:
                    logger.debug("List nesting level %d clamped to %d", logical_level, max_depth)
                    clamped_logged = True

                # Update numbering counters based on visual level hierarchy
                # (numbering uses visual_level, but tracks full logical structure)
                if visual_level <= last_visual_level:
                    # Reset deeper visual levels when going back up
                    for l in range(visual_level + 1, max_depth + 1):
                        if l in number_counters:
                            number_counters[l] = 0

                # Initialize counter for visual level if needed
                if visual_level not in number_counters:
                    number_counters[visual_level] = 0

                # Increment counter for current visual level
                number_counters[visual_level] += 1
                last_visual_level = visual_level

                # Build hierarchical number string (1., 1.1., 1.1.1.) up to visual level
                number_parts = []
                for l in range(1, visual_level + 1):
                    if l not in number_counters:
                        number_counters[l] = 1
                    number_parts.append(str(number_counters[l]))
                number_str = ".".join(number_parts) + "."

                # Create paragraph directly in container
                if is_in_table_cell and cell_elem is not None:
                    from docx.oxml import parse_xml
                    from docx.oxml.ns import nsdecls
                    p_xml = parse_xml(f'<w:p {nsdecls("w")}/>')
                    insert_at = cell_insertion_index if cell_insertion_index is not None else len(cell_elem)
                    cell_elem.insert(insert_at, p_xml)
                    new_p = Paragraph(p_xml, cell)  # type: ignore[arg-type]
                    cell_insertion_index = insert_at + 1
                else:
                    # BODY FLOW: insert at captured position
                    from docx.oxml import parse_xml
                    from docx.oxml.ns import nsdecls
                    p_xml = parse_xml(f'<w:p {nsdecls("w")}/>')
                    insertion_parent.insert(insertion_index, p_xml)
                    new_p = Paragraph(p_xml, doc._body)
                    insertion_index += 1

                # Set indentation using visual_level (consistent with numbering)
                new_p.paragraph_format.left_indent = Inches(indent_per_level * visual_level)
                new_p.paragraph_format.first_line_indent = Inches(hanging_indent)

                # Add number prefix as formatting-isolated run, then item text
                number_run = new_p.add_run(f"{number_str}  ")
                number_run.bold = False
                number_run.italic = False
                
                if formatted_runs:
                    for run_data in formatted_runs:
                        run = new_p.add_run(run_data.text)
                        run.bold = run_data.bold
                        run.italic = run_data.italic
                else:
                    new_p.add_run(item_text)

        elif block.type == "table":
            header = block.header or []
            rows = block.rows or []

            # Determine number of columns
            num_cols = len(header) if header else (len(rows[0]) if rows else 0)
            if num_cols <= 0:
                continue

            # If placeholder is in a table cell, first attempt a safe nested Word table.
            # If that is not possible, deterministically fall back to text-grid rendering.
            if is_in_table_cell:
                logger.info("Attempting nested Word table render inside table cell")
                success = False
                # Use the placeholder paragraph's current index as the insertion anchor
                start_index_for_nested = cell_insertion_index if cell_insertion_index is not None else (
                    list(cell_elem).index(paragraph._p) if cell_elem is not None else 0
                )
                if cell_elem is not None:
                    success = _attempt_render_nested_word_table(
                        block,
                        parent_cell_elem=cell_elem,
                        insertion_index=start_index_for_nested,
                        render_context=None,
                    )
                logger.info("Nested table render returned: %s", success)
                if success and cell_elem is not None:
                    # Remove the placeholder paragraph AFTER successful insertion
                    parent_for_para = paragraph._p.getparent()
                    if parent_for_para is not None and not placeholder_removed:
                        parent_for_para.remove(paragraph._p)
                        placeholder_removed = True

                    # Recompute insertion index from DOM so subsequent blocks follow the table subtree
                    children = list(cell_elem)
                    new_index = None
                    for idx, child in enumerate(children):
                        if child.tag.endswith("}tbl"):
                            new_index = idx + 1
                    if new_index is None:
                        new_index = len(children)
                    cell_insertion_index = new_index
                    continue

                # Nested table failed: now remove the placeholder paragraph and fall back to text-grid
                if is_in_table_cell and not placeholder_removed:
                    parent_for_para = paragraph._p.getparent()
                    if parent_for_para is not None:
                        parent_for_para.remove(paragraph._p)
                    placeholder_removed = True
                
                # Render header row as bold paragraph
                if header and cell_elem is not None:
                    from docx.oxml import parse_xml
                    from docx.oxml.ns import nsdecls
                    header_text = " | ".join(str(cell_text) if cell_text else "" for cell_text in header)
                    p_xml = parse_xml(f'<w:p {nsdecls("w")}/>')
                    insert_at = cell_insertion_index if cell_insertion_index is not None else len(cell_elem)
                    cell_elem.insert(insert_at, p_xml)
                    header_p = Paragraph(p_xml, cell)  # type: ignore[arg-type]
                    cell_insertion_index = insert_at + 1
                    header_run = header_p.add_run(header_text)
                    header_run.bold = True
                
                # Render separator row (if present in markdown, typically after header)
                # Note: Markdown tables don't have explicit separator rows in the block structure,
                # but we add a visual separator for clarity
                if header and rows and cell_elem is not None:
                    from docx.oxml import parse_xml
                    from docx.oxml.ns import nsdecls
                    p_xml = parse_xml(f'<w:p {nsdecls("w")}/>')
                    insert_at = cell_insertion_index if cell_insertion_index is not None else len(cell_elem)
                    cell_elem.insert(insert_at, p_xml)
                    separator_p = Paragraph(p_xml, cell)  # type: ignore[arg-type]
                    cell_insertion_index = insert_at + 1
                    separator_p.add_run("------------------------")
                
                # Render data rows as paragraphs
                if cell_elem is not None:
                    from docx.oxml import parse_xml
                    from docx.oxml.ns import nsdecls
                    for row in rows:
                        row_text = " | ".join(str(cell_text) if cell_text else "" for cell_text in row)
                        p_xml = parse_xml(f'<w:p {nsdecls("w")}/>')
                        insert_at = cell_insertion_index if cell_insertion_index is not None else len(cell_elem)
                        cell_elem.insert(insert_at, p_xml)
                        row_p = Paragraph(p_xml, cell)  # type: ignore[arg-type]
                        cell_insertion_index = insert_at + 1
                        row_p.add_run(row_text)
                
                continue

            # Create table at the exact placeholder position (BODY FLOW) or append (fallback)
            # NOTE:
            # BODY FLOW tables are constructed manually at the XML level,
            # while nested tables are built via python-docx in a scratch document.
            # This asymmetry is intentional and currently safe, but must be preserved
            # unless both paths are refactored together.
            if insertion_parent is not None and insertion_index is not None:
                # BODY FLOW: Insert table at captured position
                from docx.oxml import parse_xml
                from docx.oxml.ns import nsdecls, qn
                from docx.table import Table
                
                # Get namespace map from parent element to ensure compatibility
                parent_nsmap = insertion_parent.nsmap if hasattr(insertion_parent, 'nsmap') else {}
                w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                
                # Create table XML structure with proper namespaces
                # Use the parent's namespace context
                tbl_elem = insertion_parent.makeelement(qn('w:tbl'))
                
                # Add table properties
                tbl_pr = insertion_parent.makeelement(qn('w:tblPr'))
                tbl_elem.append(tbl_pr)
                
                # Add table grid
                tbl_grid = insertion_parent.makeelement(qn('w:tblGrid'))
                for _ in range(num_cols):
                    grid_col = insertion_parent.makeelement(qn('w:gridCol'))
                    tbl_grid.append(grid_col)
                tbl_elem.append(tbl_grid)
                
                # Add header row if present
                if header:
                    tr = insertion_parent.makeelement(qn('w:tr'))
                    for cell_text in header[:num_cols]:
                        tc = insertion_parent.makeelement(qn('w:tc'))
                        p = insertion_parent.makeelement(qn('w:p'))
                        r = insertion_parent.makeelement(qn('w:r'))
                        t = insertion_parent.makeelement(qn('w:t'))
                        t.text = str(cell_text) if cell_text else ""
                        r.append(t)
                        p.append(r)
                        tc.append(p)
                        tr.append(tc)
                    tbl_elem.append(tr)
                
                # Add data rows
                for row in rows:
                    tr = insertion_parent.makeelement(qn('w:tr'))
                    for cell_text in row[:num_cols]:
                        tc = insertion_parent.makeelement(qn('w:tc'))
                        p = insertion_parent.makeelement(qn('w:p'))
                        r = insertion_parent.makeelement(qn('w:r'))
                        t = insertion_parent.makeelement(qn('w:t'))
                        t.text = str(cell_text) if cell_text else ""
                        r.append(t)
                        p.append(r)
                        tc.append(p)
                        tr.append(tc)
                    tbl_elem.append(tr)
                
                # Insert at captured position
                insertion_parent.insert(insertion_index, tbl_elem)
                insertion_index += 1
                
                # Wrap as Table object for style application
                table = Table(tbl_elem, doc._body)
                
                # Apply default table formatting (style first, then borders fallback)
                _apply_default_table_formatting(tbl_elem, doc._body)
            else:
                # Fallback: Create table at end of document (should not happen in normal flow)
                table = doc.add_table(rows=0, cols=num_cols)
                
                # Apply default table formatting (style first, then borders fallback)
                _apply_default_table_formatting(table._tbl, doc._body)

                if header:
                    hdr_cells = table.add_row().cells
                    for idx, cell_text in enumerate(header):
                        if idx < num_cols:
                            hdr_cells[idx].text = cell_text or ""

                for row in rows:
                    row_cells = table.add_row().cells
                    for idx, cell_text in enumerate(row):
                        if idx < num_cols:
                            row_cells[idx].text = cell_text or ""


def _build_scalar_mapping(req: WordExportRequest) -> Dict[str, Optional[str]]:
    """
    Build mapping from scalar placeholder keys to values from the request.
    Keys are without braces (e.g. 'document_id'), values are raw strings.
    """
    return req.scalar_fields


def _build_plaintext_blocks(content: str) -> List[MarkdownBlock]:
    """
    Convert plain text into paragraph-only blocks.
    Does not infer bullets or tables.
    """
    if not content or not content.strip():
        return []
    parts = [p.strip() for p in content.split("\n\n") if p.strip()]
    if not parts:
        parts = [content.strip()]
    return [MarkdownBlock(type="paragraph", text=p) for p in parts]


def export_to_word(
    template_path: Path,
    request: WordExportRequest,
    markdown_mode: bool,
    output_path: Path,
    config: ListRenderConfig = ListRenderConfig()
) -> Dict[str, Any]:
    """
    Populate a Word template with scalar and block content from an existing
    summary and save the result to `output_path`.

    This function is strictly non-generative: it does not call any LLMs and
    does not alter wording, it only maps existing content into the DOCX
    structure.

    Also saves a combined markdown file (`exported_markdown_content.md`) containing
    all non-empty block contents separated by `---` lines. Empty or null content blocks are excluded.

    Returns:
        Dictionary with:
        - "word_file_path": Path to the generated DOCX file
        - "markdown_files": List containing a single path to the combined markdown file
          (`exported_markdown_content.md`), or empty list if all blocks are empty
    """
    if not template_path.exists():
        raise FileNotFoundError(f"Word template not found at: {template_path}")

    # Ensure output directory exists
    output_dir = output_path.parent
    ensure_output_dir_exists(output_dir)

    doc = Document(str(template_path))

    # 1) Replace scalar placeholders
    scalar_mapping = _build_scalar_mapping(request)
    replace_scalar_placeholders(doc, scalar_mapping)

    # 2) Replace block placeholders with structured content
    # Enforce placeholder provenance: only replace what's explicitly in request.block_fields
    # Build authoritative placeholder registry
    block_placeholders: Dict[str, Optional[str]] = {
        f"{{{{{key}}}}}": value for key, value in request.block_fields.items()
    }
    allowed_block_tokens = set(block_placeholders.keys())

    # Accumulate markdown content for all non-empty blocks
    combined_markdown_parts: List[str] = []

    # GLOBAL FREEZE: Collect ALL block placeholder targets in a single read-only traversal
    # This prevents missed replacements and non-deterministic behavior from document mutation
    frozen_block_targets: Dict[str, List[Tuple[Paragraph, Region, Container]]] = {}
    for token in allowed_block_tokens:
        frozen_block_targets[token] = []
    
    # Single read-only traversal to collect all targets
    for location_data, region, container in _iter_all_locations(doc):
        # Block expansion is only allowed in BODY FLOW and BODY TABLE
        if not _is_block_expansion_allowed(region, container):
            # Check if this location contains any block token (for warning only)
            if container == Container.TEXTBOX:
                # For textboxes, check XML runs
                run_elems = location_data
                location_text = "".join(_get_text_from_xml_run(run) for run in run_elems)
            else:
                # For paragraphs, check paragraph text
                paragraph = location_data
                if not paragraph.runs:
                    continue
                location_text = "".join(run.text or "" for run in paragraph.runs)
            
            # Check all allowed tokens for warnings
            for token in allowed_block_tokens:
                if token in location_text:
                    logger.warning(
                        f"Block placeholder {token} found inside TEXTBOX ({region.value}). "
                        "Textboxes are scalar-only. Leaving unchanged."
                    )
            continue
        
        # Only process BODY FLOW and BODY TABLE from here
        paragraph = location_data  # Must be a Paragraph for allowed locations
        if not paragraph.runs:
            continue
        para_text = "".join(run.text or "" for run in paragraph.runs)
        
        # Check all allowed tokens and collect matches
        for token in allowed_block_tokens:
            if token in para_text:
                frozen_block_targets[token].append((paragraph, region, container))

    # Now process each token using frozen targets (document mutations happen here)
    for token, content in block_placeholders.items():
        if content and content.strip():
            combined_markdown_parts.append(content)

        # Get frozen targets for this token
        frozen_targets = frozen_block_targets.get(token, [])
        
        # Invariant: log if placeholder appears multiple times
        if len(frozen_targets) > 1:
            logger.debug(f"Block placeholder {token} found in {len(frozen_targets)} locations")
        
        # Process frozen targets (document mutations happen here)
        for paragraph, region, container in frozen_targets:
            # Verify paragraph still exists and contains token (defensive check)
            try:
                if not paragraph.runs:
                    continue
                para_text = "".join(run.text or "" for run in paragraph.runs)
                if token not in para_text:
                    logger.debug(f"Block placeholder {token} no longer found in paragraph (may have been removed)")
                    continue
            except (AttributeError, RuntimeError):
                # Paragraph may have been removed by previous replacement
                logger.debug(f"Paragraph no longer accessible (may have been removed)")
                continue

            if not content or not content.strip():
                # Empty content: just remove the placeholder token and leave section empty
                logger.debug(f"Block placeholder {token} has empty content, removing placeholder only")
                replace_block_placeholder_with_content(doc, paragraph, token, blocks=None, config=config)
                continue

            if markdown_mode:
                blocks = parse_markdown_to_blocks(content)
            else:
                blocks = _build_plaintext_blocks(content)

            replace_block_placeholder_with_content(doc, paragraph, token, blocks=blocks, config=config)

    # 3) Save combined markdown (single file) alongside the DOCX
    saved_markdown_files: List[str] = []
    if combined_markdown_parts:
        combined_markdown = "\n\n---\n\n".join(combined_markdown_parts)
        markdown_file = output_dir / "exported_markdown_content.md"
        try:
            with open(markdown_file, "w", encoding="utf-8") as f:
                f.write(combined_markdown)
            saved_markdown_files.append(str(markdown_file))
            logger.info("Saved combined markdown content to %s", markdown_file)
        except Exception as e:
            logger.warning("Failed to save combined markdown file %s: %s", markdown_file, str(e))

    # 4) Save populated document
    doc.save(str(output_path))
    logger.info("Exported Word summary to %s", output_path)
    if saved_markdown_files:
        logger.info("Saved %d markdown content file(s): %s", len(saved_markdown_files), ", ".join(saved_markdown_files))
    
    return {
        "word_file_path": str(output_path),
        "markdown_files": saved_markdown_files,
    }
