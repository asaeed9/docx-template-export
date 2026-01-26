"""
Data models for deterministic DOCX export.

This module provides Pydantic models and dataclasses that define the structure
of export requests, configuration, summaries, and internal data representations.

Key Models:
- WordExportRequest: Request payload for exporting content to Word templates
- ExportSummary: Structured summary of export execution (observability)
- ListRenderConfig: Configuration for list rendering behavior
- Markdown tree node classes: Internal representation of parsed markdown structure

The models enforce type safety and validation while maintaining flexibility
for arbitrary placeholder names and content structures.
"""
