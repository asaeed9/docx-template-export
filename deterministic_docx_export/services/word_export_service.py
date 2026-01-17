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
from pathlib import Path
from typing import List, Optional, Dict, Any, Iterable

from docx import Document  # type: ignore
from docx.document import Document as DocumentType  # type: ignore
from docx.text.paragraph import Paragraph  # type: ignore
from docx.shared import Inches  # type: ignore

from markdown_it import MarkdownIt  # type: ignore

from deterministic_docx_export.models.export_config import ListRenderConfig
from deterministic_docx_export.models.export_models import WordExportRequest
from deterministic_docx_export.services.output_path import get_output_dir_for_document, ensure_output_dir_exists

logger = logging.getLogger(__name__)


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

            # Paragraphs
            if tok.type == "paragraph_open":
                if i + 1 < len(tokens) and tokens[i + 1].type == "inline":
                    inline_token = tokens[i + 1]
                    para_text = inline_token.content or ""
                    if para_text.strip():
                        formatted_runs = _parse_inline_formatting(inline_token)
                        blocks.append(
                            MarkdownBlock(
                                type="paragraph",
                                text=para_text,  # Keep for fallback
                                formatted_runs=formatted_runs if formatted_runs else None,
                            )
                        )
                # Skip until paragraph_close
                while i < len(tokens) and tokens[i].type != "paragraph_close":
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
                
                def parse_ordered_list(start_idx: int, base_level: int) -> int:
                    """
                    Parse an ordered list starting at start_idx.
                    Returns index after the matching ordered_list_close.
                    base_level: nesting level (1 for top-level, 2 for nested, etc.)
                    """
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


def _iter_all_paragraphs(doc: DocumentType) -> Iterable[Paragraph]:
    """
    Yield all paragraphs in the document, including paragraphs inside table cells.
    """
    for p in doc.paragraphs:
        yield p

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


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

    # Build the concatenated text
    full_text = "".join(run.text or "" for run in runs)
    if not full_text or placeholder not in full_text:
        return

    while True:
        idx = full_text.find(placeholder)
        if idx == -1:
            break
        end = idx + len(placeholder)

        new_texts: List[str] = []
        current_start = 0
        inserted = False

        for run in runs:
            original = run.text or ""
            run_len = len(original)
            run_start = current_start
            run_end = current_start + run_len

            # No overlap with placeholder
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

        # Apply updates back to runs
        for run, new_text in zip(runs, new_texts):
            run.text = new_text

        # Recompute combined text for potential further replacements
        full_text = "".join(run.text or "" for run in runs)
        if placeholder not in full_text:
            break


def replace_scalar_placeholders(doc: DocumentType, mapping: Dict[str, Optional[str]]) -> None:
    """
    Replace scalar placeholders (e.g. {{document_id}}) in all paragraphs and table cells.

    This function only replaces inline text and never inserts or removes paragraphs.
    """
    # Build concrete placeholder -> replacement mapping
    concrete: Dict[str, str] = {}
    for key, value in mapping.items():
        token = f"{{{{{key}}}}}"  # e.g. document_id -> {{document_id}}
        concrete[token] = value or ""

    if not concrete:
        return

    for paragraph in _iter_all_paragraphs(doc):
        if not paragraph.runs:
            continue
        para_text = "".join(run.text or "" for run in paragraph.runs)
        if not para_text:
            continue

        for token, replacement in concrete.items():
            if token in para_text:
                _replace_text_in_runs(paragraph.runs, token, replacement)
                # Update cached text for subsequent tokens
                para_text = "".join(run.text or "" for run in paragraph.runs)


def _insert_after(anchor, new_element) -> None:
    """
    Insert `new_element` in the document XML immediately after `anchor`.
    """
    parent = anchor.getparent()
    if parent is None:
        return
    index = parent.index(anchor)
    parent.insert(index + 1, new_element)


def _insert_paragraph_after(doc: DocumentType, paragraph: Paragraph, text: str = "", style: Optional[str] = None, formatted_runs: Optional[List[FormattedRun]] = None) -> Paragraph:
    """
    Create a new paragraph after the given paragraph, with optional style and text.
    If formatted_runs is provided, uses those instead of plain text.
    """
    new_p = doc.add_paragraph()
    if style:
        new_p.style = style
    
    if formatted_runs:
        # Add formatted runs
        for run_data in formatted_runs:
            run = new_p.add_run(run_data.text)
            run.bold = run_data.bold
            run.italic = run_data.italic
    elif text:
        new_p.text = text

    _insert_after(paragraph._p, new_p._p)  # type: ignore[attr-defined]
    return new_p


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
    # Defensive guard: ensure config is never None
    if config is None:
        raise ValueError("config must not be None")
    
    # Remove the placeholder token text from this paragraph
    _replace_text_in_runs(paragraph.runs, placeholder_token, "")

    # If no content blocks, we leave the section visually empty.
    if not blocks:
        return

    anchor = paragraph._p  # xml element used for positioning subsequent content
    used_existing_para = False

    for block in blocks:
        if block.type == "heading":
            text = block.text or ""
            style = _heading_style_for_level(block.level)

            if not used_existing_para:
                # Reuse the existing paragraph for the first block
                _set_paragraph_formatted_text(paragraph, block.formatted_runs, text)
                paragraph.style = style
                anchor = paragraph._p
                used_existing_para = True
            else:
                new_p = _insert_paragraph_after(doc, Paragraph(anchor, paragraph._parent), text=text, style=style, formatted_runs=block.formatted_runs)  # type: ignore[arg-type]
                anchor = new_p._p

        elif block.type == "paragraph":
            text = block.text or ""
            if not used_existing_para:
                _set_paragraph_formatted_text(paragraph, block.formatted_runs, text)
                # Keep existing style (body style from template)
                anchor = paragraph._p
                used_existing_para = True
            else:
                new_p = _insert_paragraph_after(doc, Paragraph(anchor, paragraph._parent), text=text, formatted_runs=block.formatted_runs)  # type: ignore[arg-type]
                anchor = new_p._p

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

                # Create paragraph and insert after anchor
                new_p = doc.add_paragraph()
                _insert_after(anchor, new_p._p)

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

                anchor = new_p._p
                if idx == 0:
                    used_existing_para = True

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

                # Create paragraph and insert after anchor
                new_p = doc.add_paragraph()
                _insert_after(anchor, new_p._p)

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

                anchor = new_p._p
                if idx == 0:
                    used_existing_para = True

        elif block.type == "table":
            header = block.header or []
            rows = block.rows or []

            # Determine number of columns
            num_cols = len(header) if header else (len(rows[0]) if rows else 0)
            if num_cols <= 0:
                continue

            table = doc.add_table(rows=0, cols=num_cols)
            
            # Apply table style - try "Light Grid" first (Table Grid Light design)
            # Falls back to other grid styles if not available
            try:
                table.style = 'Light Grid'
            except Exception:
                try:
                    table.style = 'Light Grid Accent 1'
                except Exception:
                    try:
                        table.style = 'Table Grid'
                    except Exception:
                        # If no style works, leave it unstyled (content fidelity is primary)
                        pass

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

            # Position the table immediately after the current anchor
            _insert_after(anchor, table._tbl)  # type: ignore[attr-defined]
            anchor = table._tbl


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
    block_placeholders: Dict[str, Optional[str]] = {
        f"{{{{{key}}}}}": value for key, value in request.block_fields.items()
    }

    # Accumulate markdown content for all non-empty blocks
    combined_markdown_parts: List[str] = []

    for token, content in block_placeholders.items():
        if content and content.strip():
            combined_markdown_parts.append(content)

        # Find each paragraph containing this token
        for paragraph in _iter_all_paragraphs(doc):
            para_text = "".join(run.text or "" for run in paragraph.runs)
            if token not in para_text:
                continue

            if not content or not content.strip():
                # Empty content: just remove the placeholder token and leave section empty
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
