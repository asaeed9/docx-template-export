# Changelog

**Purpose**: This file provides a technical changelog for developers and maintainers. It documents detailed changes, API modifications, breaking changes, bug fixes, and internal improvements. For user-friendly summaries, see the "What's New" sections in README.md.

**Scope**: This changelog focuses on:
- Major features and enhancements (with technical details)
- API changes (function signatures, return values, new parameters)
- Breaking changes (if any)
- Bug fixes (with technical context and root causes)
- Internal improvements (refactoring, code organization)
- Environment variables and configuration changes
- Testing infrastructure changes
- Documentation additions (new docs, major updates)

---

## v0.3.1

### Package Rename

- **Package renamed**: The package has been renamed from `deterministic-docx-export` to `docx-template-export`
- **Backward compatibility**: The old package name (`deterministic-docx-export`) remains functional via a compatibility shim that emits deprecation warnings
- **Import changes**: New code should use `from docx_template_export import ...` instead of `from deterministic_docx_export import ...`
- **No functional changes**: All functionality, APIs, and behavior remain identical; only the package name has changed
- **Version consistency**: Both package names report the same version (0.3.1)

### Bug Fixes

- **Cross-platform table formatting**: Tables now always have explicit XML borders defined, ensuring consistent rendering on Windows Word. Table style "TableGrid" is preferred, and explicit borders are applied even when styles succeed (Windows Word requires explicit XML borders for consistent display).

### Backward Compatibility

- Fully backward compatible; all changes are additive
- No impact on output, control flow, or rendering behavior
- All existing tests pass unchanged
- Old package name imports continue to work (with deprecation warnings)

## v0.3.0

### Major Feature: Tree-Based Rendering Engine

**Primary Enhancement**: This release introduces a new tree-based rendering engine as the default rendering path, replacing block-based rendering as the primary mechanism while maintaining block-based rendering as a fallback.

- **Tree-Based Renderer**: New experimental tree-based rendering engine (`renderers/tree_renderer.py`) that attempts to preserve exact order and nesting of markdown elements (lists, list items, tables, table cells)
- **Enabled by Default**: Tree renderer is enabled by default in v0.3.0. Controlled via `USE_MARKDOWN_TREE_RENDERER` environment variable (default: `True`)
- **Tree Intermediate Representation**: New `models/markdown_tree.py` module providing immutable tree structure (`DocumentNode`, `HeadingNode`, `ParagraphNode`, `ListNode`, `ListItemNode`, `TableNode`, `TableRowNode`, `TableCellNode`)
- **Enhanced Structure Preservation**: Tree renderer attempts to preserve semantic structure without flattening, allowing for more accurate rendering of complex nested structures
- **Graceful Fallback**: Tree parsing failures automatically fall back to block-based renderer via `parse_markdown_to_tree()` returning `None` on failure
- **Safety Parity**: Tree renderer provides similar safety guarantees as block-based renderer:
  - Atomic operations for table insertion (scratch → clone → insert pattern)
  - Freeze → validate → mutate pattern for placeholder replacement
  - Tree depth and node count validation before rendering (`_validate_tree_safety()`)
- **Block-Based Renderer**: Remains available as the fallback path when tree renderer is disabled or encounters errors

**Implementation Details**:
- Tree parsing: `parse_markdown_to_tree()` function in `word_export_service.py`
- Tree rendering: `render_markdown_tree_to_docx()` function in `renderers/tree_renderer.py`
- Environment variable: `USE_MARKDOWN_TREE_RENDERER` (default: `"true"`, can be set to `"false"` to disable)
- Tree safety limits: `MAX_TREE_DEPTH=50`, `MAX_NODE_COUNT=10000` (internal constants)

### Major Feature: Observability & Export Summary

**Primary Enhancement**: This release adds comprehensive observability features to track export execution, enabling users to audit what happened during document generation without affecting output or performance.

- **Export Execution Summary**: New `ExportSummary` dataclass (`models/export_summary.py`) that records:
  - Scalar replacements (token, region, container, value preview)
  - Block replacements (token, region, container, block types)
  - Skipped items (token, region, container, reason)
  - Fallback events (type, location, reason)
  - Warnings (message, context)
- **Structured Logging**: Emit structured logs aligned with export summary, including successful replacements, skipped placeholders, fallback events, and safety constraint warnings
- **Automatic Log File Saving**: Structured logs are automatically saved to `export_trace_{filename}.log` files in the same timestamped directory as JSON summaries
- **Enable/Disable Control**: New `enable_export_trace` flag in `WordExportRequest` (default: `True`) to gate summary collection and structured logging. When disabled, incurs near-zero overhead
- **Automatic JSON Summary Saving**: Export summaries are automatically saved to timestamped JSON files (`export_summary_{filename}.json`) in a configurable log directory structure
- **Environment Variable Support**: 
  - `EXPORT_LOG_DIR`: Customize the base directory for export logs. Logs are organized in timestamped subdirectories (e.g., `logs/20240124_143022/`)
  - Both JSON summary and structured log files are saved to the same timestamped directory per export run
- **Safety Invariant Documentation**: Added inline comments explaining critical design decisions and safety guarantees:
  - Why block placeholders must occupy entire paragraphs (block markdown cannot be mixed with text in the same paragraph)
  - Why textboxes only support scalar replacement
  - Why ambiguous placeholders are skipped
  - Why the tree renderer is feature-flagged
- **Tree Renderer Observability**: Detect and log (without changing behavior) tree parsing failures, nested tables inside cells, and mixed ordered/bullet lists when tree-based renderer is enabled via `_observe_tree_structure()` function

### API Changes

- **Return Value Enhancement**: `export_to_word()` function signature unchanged, but return dictionary now includes:
  - `export_summary`: Dictionary representation of `ExportSummary` (when `enable_export_trace=True`)
  - `export_summary_log_path`: Path to the saved JSON log file (when `enable_export_trace=True`)
  - Structured log file (`export_trace_{filename}.log`) is automatically created in the same directory but path is not returned (can be derived from `export_summary_log_path`)
- **Request Model Enhancement**: `WordExportRequest` class (`models/export_models.py`) now includes:
  - `enable_export_trace`: Boolean field (default: `True`) to enable/disable observability features
- **New Public Functions**: 
  - `parse_markdown_to_tree()`: Build tree-based IR from markdown (feature-flagged, experimental)
  - `_observe_tree_structure()`: Detect and log tree structure events (internal, called during export)
- **New Modules**:
  - `renderers/tree_renderer.py`: Tree-based rendering implementation
  - `models/markdown_tree.py`: Tree intermediate representation (IR) data structures

### Environment Variables

- **`OUTPUT_DIRECTORY`**: Base directory for output files (new in v0.3.0)
  - Default: Not set (relative paths resolved relative to current working directory)
  - If set: Relative `output_path` parameters are resolved relative to this directory
  - If `output_path` is absolute, this variable is ignored (absolute paths used as-is)
  - Useful for centralizing all exported documents in a single location
  - Read in `word_export_service.py` via `os.getenv("OUTPUT_DIRECTORY")`
  - Example: `OUTPUT_DIRECTORY=/var/exports` → `output_path=Path("doc.docx")` becomes `/var/exports/doc.docx`
- **`USE_MARKDOWN_TREE_RENDERER`**: Control tree-based rendering engine (new in v0.3.0)
  - Default: `"true"` (tree renderer enabled by default)
  - Set to `"false"`, `"0"`, or `"no"` to disable and use block-based rendering
  - Read in `word_export_service.py` via `os.getenv("USE_MARKDOWN_TREE_RENDERER", "true")`
  - When enabled, block placeholders are processed using tree-based parser and renderer
  - When disabled, uses block-based renderer (legacy behavior)
- **`EXPORT_LOG_DIR`**: Customize base directory for export logs (new in v0.3.0)
  - If set: Logs saved to `{EXPORT_LOG_DIR}/logs/{timestamp}/`
  - If not set: Logs saved to `{output_dir}/logs/{timestamp}/`
  - Each export run creates one timestamped subdirectory containing both JSON summary and structured log file
  - Read in `word_export_service.py` via `os.getenv("EXPORT_LOG_DIR")`

### Documentation

- **OBSERVABILITY.md**: Comprehensive guide covering:
  - Quick answers on where to find summaries and logs
  - Log directory structure and environment variables
  - Complete code examples
  - Summary structure documentation
  - Enable/disable instructions
  - Automatic log file saving (both JSON and structured logs)
- **README.md**: Updated with:
  - Tree renderer as major feature in v0.3.0
  - Observability section with environment variable documentation
  - Examples showing both JSON summary and structured log files
  - Tree renderer configuration instructions
- **Inline Comments**: Added safety invariant documentation throughout the codebase explaining design rationale
- **Module Docstrings**: Enhanced all module, class, and function docstrings with detailed descriptions, safety guarantees, and usage examples

### Internal Improvements

- **Code Documentation**: Comprehensive docstring review and enhancement across all modules:
  - Package-level docstrings in `__init__.py` files
  - Enhanced class docstrings with detailed attribute descriptions
  - Enhanced function docstrings with purpose, arguments, returns, safety guarantees, and design rationale
  - Added inline comments explaining "WHY" decisions (safety invariants, design choices)
- **Documentation Language**: Updated all docstrings to use conservative, safety-aware language:
  - Replaced absolute claims ("always", "guarantees", "preserves") with precise language ("attempts to", "where structurally safe", "subject to DOCX constraints")
  - Emphasized determinism over completeness, safety over aggressive rendering
  - Clarified experimental/feature-flagged nature of tree renderer
- **Example Consolidation**: Merged `example_view_summary.py` into `example_usage.py` as Example 4, demonstrating observability features including automatic log file saving

### Backward Compatibility

- Fully backward compatible; all changes are additive
- Default behavior enables tree renderer (`USE_MARKDOWN_TREE_RENDERER=true`) and observability (`enable_export_trace=True`)
- No impact on output, control flow, or rendering behavior when defaults are used
- Block-based renderer remains available as fallback (automatic on tree parsing failure, or explicit via environment variable)
- All existing tests pass unchanged (108/110 tests passing; 2 pre-existing failures unrelated to v0.3.0 changes)

## v0.2.1

### Critical Bug Fixes: Text Loss Prevention

**Primary Issue**: This release fixes critical bugs that caused **markdown paragraphs to be silently lost** during export, violating the library's loss-less text guarantee.

- **Continuation paragraph preservation**: Fixed bug where multi-paragraph list items lost continuation paragraphs. Continuation paragraphs are now preserved with `_list_continuation_level` semantic contract and rendered with proper indentation
- **Nested bullet list text loss**: Fixed critical bug where nested bullet lists inside ordered lists were completely skipped, causing guaranteed text loss. Implemented `parse_nested_bullet_for_ordered()` to collect ALL paragraphs within nested bullet items, not just the first paragraph
- **Free text before lists**: Fixed parsing bug where top-level paragraphs appearing before lists could be incorrectly merged or lost
- **Unsafe paragraph clearing**: Replaced `paragraph.clear()` with run-level text clearing to preserve paragraph-level styles, spacing, numbering, and template formatting. This prevents potential text loss and formatting corruption
- **Placeholder detection failures**: Fixed bug where placeholders spanning multiple XML runs or containing line breaks (`<w:br/>`/`<w:cr/>`) could be missed during scanning. All placeholder detection now uses consistent XML-based text extraction (`_get_text_from_xml_run()`) throughout the codebase

### Supporting Fixes

- **List spacing fix**: Neutralized Word's default paragraph spacing for list items (bullets and numbered) and continuation paragraphs. Spacing now comes exclusively from markdown (`\n\n`) via `paragraph_spacing_before_pt`, not from Word's default paragraph styles. This ensures visual consistency and prevents spacing-related rendering issues

### Testing & Quality

- **Production-grade test suite**: Added comprehensive test suite with 28 tests organized by capability (scalar replacement, block markdown, lists, tables, textboxes, documented limitations, negative cases, golden regression)
- **Text fidelity enforcement**: All tests enforce Invariant 1 (Text Fidelity) using `assert_text_fidelity()` helper, ensuring no text loss. This test suite was specifically designed to catch the text loss bugs that were fixed in this release
- **Text loss regression tests**: Added comprehensive regression tests (`TestTextLossRegression`) that validate all known text-loss edge cases in a single test scenario, ensuring continuation paragraphs, nested lists, and free text are never lost
- **Test artifacts**: All tests write artifacts (input.json, rendered_markdown.md, output.docx) for human inspection, enabling visual verification that no text is missing
- **Test summary**: Added `generate_summary.py` script that generates `TEST_RESULTS_SUMMARY.md` with detailed test descriptions and outcomes

### Documentation

- **Enhanced docstrings**: Added explicit design contracts for list continuation paragraphs and block placeholder rules
- **Test documentation**: Comprehensive test descriptions explaining what each test validates and which invariants it enforces

### Backward Compatibility

- No breaking API changes; fully backward compatible
- All existing functionality preserved
- List spacing fix improves visual consistency without changing behavior

## v0.2.0

### Major Enhancements

- Safe nested Word table rendering with atomic commit and deterministic fallback
- Deterministic table formatting using Word styles with XML border fallback
- Multi-run�safe scalar placeholder replacement in headers, footers, and textboxes

### Safety & Guardrails

- Atomic block placeholder expansion (freeze → validate → mutate)
- Strict BODY-only block expansion rules
- Guaranteed structural validity of table cells after replacement

### Internal Improvements

- Removal of unused legacy helpers and added invariant documentation
- Improved code organization with separated table formatting responsibilities
- Simplified cell detection logic for better maintainability

### Backward Compatibility

- No breaking API changes; fully backward compatible
- All existing functionality preserved with improved safety guarantees
