# Copyright ? 2024 Ahsan Saeed
# Licensed under the Apache License, Version 2.0
# See LICENSE and NOTICE files for details.

"""
Tree-based intermediate representation (IR) for markdown parsing.

This module provides an immutable tree structure that attempts to preserve order
and nesting of markdown elements (lists, list items, tables, table cells) where
structurally safe. This tree representation is used by the tree-based renderer,
which is an experimental, feature-flagged component that runs in parallel to
the block-based renderer.

The tree IR attempts to preserve semantic structure without flattening or losing
nesting information where supported. This may allow for more accurate rendering
of complex nested structures (e.g., lists within lists, tables within list items),
subject to DOCX constraints.

Tree Structure:
    The tree is a hierarchical structure of BlockNode subclasses:
    - DocumentNode: Root node containing all top-level blocks
    - HeadingNode: Markdown headings (H1-H6)
    - ParagraphNode: Plain paragraphs with formatted text
    - ListNode: Bullet or ordered lists containing ListItemNode children
    - ListItemNode: Individual list items that can contain any block-level children
    - TableNode: Tables containing TableRowNode children
    - TableRowNode: Table rows containing TableCellNode children
    - TableCellNode: Table cells that can contain any block-level children

Key Design Decisions:
    - Immutable structure: Nodes are dataclasses that are not modified after creation
    - Recursive nesting: ListItemNode and TableCellNode can contain any BlockNode,
      allowing arbitrary nesting depth
    - Formatted runs: Text content is stored as FormattedRun objects that preserve
      bold/italic formatting information
    - Type safety: Literal types ensure list kind is either "bullet" or "ordered"

Relationship to Block-Based Renderer:
    The tree IR is parallel to the existing MarkdownBlock representation and does
    not replace it. The block-based renderer remains the default and primary path.
    The tree renderer is feature-flagged and used only when explicitly enabled or
    when the block renderer encounters structures it cannot handle.

Safety Guarantees:
    - Tree depth and node count are validated before rendering to attempt to prevent stack overflow
    - Tree parsing failures gracefully fall back to block-based renderer
    - Tree operations attempt to preserve document structure integrity where structurally safe
"""
from dataclasses import dataclass, field
from typing import List, Optional, Literal

# Reuse FormattedRun from word_export_service
# Note: This import is safe because word_export_service only imports from this module
# inside functions (lazy imports), not at module level
from docx_template_export.services.word_export_service import FormattedRun


@dataclass
class BlockNode:
    """
    Base class for all markdown tree nodes.
    
    All nodes in the markdown tree inherit from this base class. The base class
    itself has no attributes; it serves as a type marker for the type system
    and allows polymorphic handling of different node types.
    """
    pass


@dataclass
class DocumentNode(BlockNode):
    """
    Root node representing the entire document.
    
    This is the top-level node of the markdown tree. It contains all top-level
    blocks (headings, paragraphs, lists, tables) as its children. There is
    exactly one DocumentNode per parsed markdown document.
    
    Attributes:
        children: List of top-level BlockNode instances representing the document
            structure. Order attempts to be preserved from the original markdown.
    """
    children: List[BlockNode] = field(default_factory=list)


@dataclass
class HeadingNode(BlockNode):
    """
    Heading node representing a markdown heading (H1-H6).
    
    Headings preserve their level (1-6) and formatted text content. Inline
    formatting (bold, italic) is preserved through FormattedRun objects.
    
    Attributes:
        level: Heading level (1-6), where 1 is the highest level (largest).
        runs: List of FormattedRun objects representing the heading text with
            formatting preserved. If empty, the heading has no text content.
    """
    level: int
    runs: List[FormattedRun] = field(default_factory=list)


@dataclass
class ParagraphNode(BlockNode):
    """
    Paragraph node representing a plain text paragraph.
    
    Paragraphs contain formatted text content. Inline formatting (bold, italic)
    is preserved through FormattedRun objects.
    
    Attributes:
        runs: List of FormattedRun objects representing the paragraph text with
            formatting preserved. If empty, the paragraph is empty.
    """
    runs: List[FormattedRun] = field(default_factory=list)


@dataclass
class ListNode(BlockNode):
    """
    List node representing a bullet or ordered list.
    
    Lists contain list items as children. The list type (bullet or ordered)
    is determined by the `kind` attribute. All items in a list share the same
    type (bullet lists cannot contain ordered sub-lists directly, though
    nested lists can be created through ListItemNode children).
    
    Attributes:
        kind: Type of list, either "bullet" (unordered) or "ordered" (numbered).
            This is a Literal type to ensure type safety.
        items: List of ListItemNode instances representing the list items.
            Order attempts to be preserved from the original markdown.
    """
    kind: Literal["bullet", "ordered"]
    items: List["ListItemNode"] = field(default_factory=list)


@dataclass
class ListItemNode(BlockNode):
    """
    List item node that can contain any block-level children.
    
    List items are containers that can hold any type of BlockNode as children.
    This allows for nested structures such as:
    - Lists within list items (nested lists)
    - Tables within list items
    - Paragraphs and headings within list items
    - Mixed content within a single list item
    
    Attributes:
        children: List of BlockNode instances representing the content of this
            list item. Can contain headings, paragraphs, nested lists, tables,
            or any combination thereof. Order attempts to be preserved from the original markdown.
    """
    children: List[BlockNode] = field(default_factory=list)


@dataclass
class TableNode(BlockNode):
    """
    Table node representing a markdown table.
    
    Tables contain rows, which in turn contain cells. The table structure
    attempts to be preserved as it appears in the markdown, including row and
    column counts, subject to DOCX constraints.
    
    Attributes:
        rows: List of TableRowNode instances representing the table rows.
            Order attempts to be preserved from the original markdown. The first row is
            typically the header row, though this is not enforced by the structure.
    """
    rows: List["TableRowNode"] = field(default_factory=list)


@dataclass
class TableRowNode(BlockNode):
    """
    Table row node representing a single row in a table.
    
    Rows contain cells as children. All rows in a table should have the same
    number of cells for proper table structure, though this is not enforced
    by the tree structure itself.
    
    Attributes:
        cells: List of TableCellNode instances representing the cells in this row.
            Order determines column position (first cell = first column, etc.).
    """
    cells: List["TableCellNode"] = field(default_factory=list)


@dataclass
class TableCellNode(BlockNode):
    """
    Table cell node that can contain any block-level children.
    
    Table cells are containers that can hold any type of BlockNode as children.
    This allows for rich content within cells, including:
    - Paragraphs and headings
    - Lists (bullets or numbered)
    - Nested tables (though these may trigger fallback rendering)
    - Mixed content
    
    Attributes:
        children: List of BlockNode instances representing the content of this
            table cell. Can contain headings, paragraphs, lists, nested tables,
            or any combination thereof. Order attempts to be preserved from the original markdown.
            
    Note:
        Nested tables within cells are detected and may trigger fallback to
        text-based rendering for safety and determinism.
    """
    children: List[BlockNode] = field(default_factory=list)
