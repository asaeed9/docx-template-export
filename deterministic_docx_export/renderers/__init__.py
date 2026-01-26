# Copyright ù 2024 Ahsan Saeed
# Licensed under the Apache License, Version 2.0
# See LICENSE and NOTICE files for details.

"""
Renderers for markdown tree to DOCX conversion.

This package provides rendering implementations that attempt to convert parsed markdown
structures into Word document elements. The tree-based renderer is an experimental,
feature-flagged component that attempts to preserve nesting and order of markdown elements
where structurally safe.

Main Renderers:
- tree_renderer: Experimental tree-based renderer that processes markdown tree structures
    (DocumentNode, ListNode, etc.) and attempts to render them to Word documents with
    order and nesting preservation where supported. This is feature-flagged and runs
    in parallel to the block-based renderer (which remains the default).

The tree renderer attempts to provide safety parity with the block-based renderer:
- Similar safety invariants (atomic operations, freeze ? validate ? mutate) where applicable
- Similar list rendering semantics (manual glyph/numbering, no Word numbering XML) where supported
- Graceful fallback to block-based renderer on parsing failures or unsupported structures
"""
