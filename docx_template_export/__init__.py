"""
Deterministic DOCX Export Engine.

A production-grade Python library for exporting structured content (markdown or plain text)
into Word (.docx) templates with deterministic rendering and template safety guarantees.

This package provides:
- Scalar placeholder replacement (e.g., {{document_id}}, {{title}}) where structurally safe
- Block placeholder replacement with structured content ({{summary}}, {{proposal}}, etc.) in BODY containers only
- Markdown parsing and conversion to Word structures (headings, lists, tables) where supported
- Deterministic, cross-platform rendering that attempts to produce identical output given identical inputs
- Text-preserving fallback: content is preserved as text when structure cannot be rendered
- Template safety: block placeholders must occupy entire paragraphs; textboxes are scalar-only

The library is strictly non-generative: it does not call any LLMs and does not modify wording.
It only maps existing content into Word document structures.

Key Design Principles:
- Determinism: Attempts to produce identical outputs given identical inputs across runs and platforms
- Safety: Block expansion is restricted to BODY containers only (headers/footers/textboxes excluded)
- Text Preservation: Content is preserved as text when structure cannot be rendered (text-preserving fallback)
- Stability: Manual list rendering (no Word numbering XML) for cross-platform consistency, subject to DOCX constraints

Main Entry Point:
    export_to_word() - Primary function for exporting content to Word templates

Version:
    The package version is automatically loaded from package metadata.
    Use __version__ to access the current version string.
"""

try:
    from importlib.metadata import version
    __version__ = version("docx-template-export")
except Exception:
    # Fallback if package is not installed (e.g., during development)
    __version__ = "0.3.1"
