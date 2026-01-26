# Deterministic DOCX Export

A standalone Python library for exporting structured content (markdown or plain text) into Word (.docx) templates. This module is **strictly non-generative** - it does not call any LLMs and does not modify wording. It only maps existing content into Word document structures.

## Author

Created and maintained by **Ahsan Saeed**.

This project is designed with a strong focus on determinism, correctness, and enterprise-grade document generation.

## What's New in v0.3.0

**Major Features: Tree-Based Rendering & Observability**

This release introduces a new tree-based rendering engine enabled by default, along with comprehensive observability features to track export execution. The library has moved from block-based rendering to tree-based rendering as the primary path, with block-based rendering remaining as a fallback.

### Key Features

#### Tree-Based Rendering (Enabled by Default)

- **Tree-Based Renderer**: New tree-based rendering engine that preserves exact order and nesting of markdown elements (lists, list items, tables, table cells)
- **Enabled by Default**: Tree renderer is enabled by default in v0.3.0. Can be disabled via `USE_MARKDOWN_TREE_RENDERER` environment variable (set to `"false"` to disable)
- **Enhanced Structure Preservation**: Tree renderer attempts to preserve semantic structure without flattening, allowing for more accurate rendering of complex nested structures
- **Graceful Fallback**: Tree parsing failures automatically fall back to block-based renderer to ensure export completion
- **Safety Parity**: Tree renderer provides similar safety guarantees as block-based renderer (atomic operations, freeze → validate → mutate)
- **Block-Based Renderer**: Remains available as the fallback path when tree renderer is disabled or encounters errors

#### Observability & Export Summary

- **Export Execution Summary**: New `ExportSummary` dataclass that records scalar replacements, block replacements, skipped items, fallback events, and warnings
- **Structured Logging**: Emit structured logs aligned with export summary, including successful replacements, skipped placeholders, fallback events, and safety constraint warnings
- **Enable/Disable Control**: New `enable_export_trace` flag in `WordExportRequest` (default: `True`) to gate summary collection and structured logging
- **Automatic Log Saving**: Export summaries are automatically saved to timestamped JSON files in a configurable log directory structure
- **Tree Renderer Observability**: Detect and log tree parsing failures, nested tables inside cells, and mixed ordered/bullet lists when tree-based renderer is enabled
- **Safety Invariant Documentation**: Added inline comments explaining critical design decisions, including why block placeholders must occupy entire paragraphs (block markdown cannot be mixed with text in the same paragraph)

### API Changes

- `export_to_word()` now returns `export_summary` (dictionary) and `export_summary_log_path` (string) when `enable_export_trace=True`
- `WordExportRequest` now includes `enable_export_trace` boolean field (default: `True`)

### Environment Variables

Three new environment variables are introduced in v0.3.0:

- **`OUTPUT_DIRECTORY`**: Base directory for output files
  - Default: Not set (relative paths resolved relative to current working directory)
  - If set, relative `output_path` parameters are resolved relative to this directory
  - Absolute `output_path` parameters are used as-is (this variable is ignored)
  - Useful for centralizing all exported documents in a single location

- **`USE_MARKDOWN_TREE_RENDERER`**: Control tree-based rendering engine (default: `True` - enabled)
  - Default: `True` - Tree renderer is enabled by default in v0.3.0
  - Set to `"false"`, `"0"`, or `"no"` to disable and use block-based rendering
  - When enabled (default), the library uses tree-based parsing and rendering for block placeholders
  - When disabled, uses block-based rendering (legacy behavior)
  - Tree renderer provides enhanced structure preservation for complex nested markdown

- **`EXPORT_LOG_DIR`**: Customize the base directory for export summary logs (optional)
  - If set: Logs saved to `{EXPORT_LOG_DIR}/logs/{timestamp}/`
  - If not set: Logs saved to `{output_dir}/logs/{timestamp}/`
  - Each export run gets its own timestamped subdirectory

The library automatically loads `.env` files if `python-dotenv` is available.

All changes are fully backward compatible with no breaking API changes. See [OBSERVABILITY.md](OBSERVABILITY.md) for detailed documentation.

## What's New in v0.2.1

**Critical Release: Text Loss Bug Fixes**

This release fixes critical bugs that caused **markdown paragraphs to be silently lost** during export. The library's core guarantee is loss-less text preservation, and these fixes ensure that guarantee is maintained.

### Primary Fixes: Text Loss Prevention

- **Continuation paragraph preservation**: Fixed bug where multi-paragraph list items (both bullet and numbered) lost their continuation paragraphs. Continuation paragraphs are now preserved and rendered with proper indentation
- **Nested bullet list text loss**: Fixed critical bug where nested bullet lists inside ordered lists were completely skipped, causing guaranteed text loss. The parser now collects ALL paragraphs within nested bullet items
- **Free text before lists**: Fixed parsing bug where top-level paragraphs appearing before lists could be incorrectly merged or lost
- **Unsafe paragraph clearing**: Replaced unsafe `paragraph.clear()` with run-level text clearing to preserve formatting and prevent potential text loss
- **Placeholder detection failures**: Fixed bug where placeholders spanning multiple XML runs or containing line breaks could be missed. All placeholder detection now uses consistent XML-based text extraction

### Supporting Improvements

- **List spacing fix**: Neutralized Word's default paragraph spacing for list items and continuation paragraphs, ensuring spacing comes from markdown (`\n\n`) rather than Word defaults
- **Production-grade test suite**: Added comprehensive test suite with 28 tests enforcing text fidelity, determinism, and observability invariants. Includes specific regression tests for all known text-loss edge cases
- **Documentation improvements**: Enhanced docstrings and added explicit design contracts for list continuation paragraphs and block placeholder rules

All changes are fully backward compatible with no breaking API changes.

## What's New in v0.2.0

This release focuses on deterministic table rendering and safety hardening:

- **Safe nested Word table rendering**: Tables can now be rendered inside table cells with atomic commit and deterministic fallback to text rendering
- **Deterministic table formatting**: Tables use Word's built-in styles (Light Grid, Medium Grid 1, Table Grid) with XML border fallback for consistent visual appearance
- **Multi-run–safe scalar replacement**: Scalar placeholders spanning multiple XML runs are safely replaced in headers, footers, and textboxes
- **Atomic block placeholder expansion**: Block placeholders use a freeze → validate → mutate model ensuring no partial mutations
- **Strict container guardrails**: Block expansion is restricted to BODY containers only, preventing unsafe operations in headers/footers/textboxes
- **Structural integrity guarantees**: Table cells maintain proper XML structure after placeholder replacement
- **Code quality improvements**: Removal of unused legacy helpers and addition of invariant documentation for maintainability

All changes are fully backward compatible with no breaking API changes.

## Features

- **Scalar placeholder replacement**: Replace simple placeholders like `{{document_id}}`, `{{title}}`, etc.
- **Block placeholder replacement**: Replace structured content blocks like `{{summary}}`, `{{proposal}}`, etc.
- **Tree-based rendering** (enabled by default): Enhanced rendering engine that attempts to preserve exact order and nesting of markdown elements (can be disabled via `USE_MARKDOWN_TREE_RENDERER` environment variable)
- **Block-based rendering** (fallback): Stable, proven rendering path that processes markdown as independent blocks (used when tree renderer is disabled or encounters errors)
- **Markdown parsing**: Convert markdown content to Word structures (headings, lists, tables)
- **Manual list rendering**: Bullet and numbered lists use deterministic glyph/number insertion (no Word numbering XML) for stable, cross-platform rendering
- **Plain text mode fallback**: Paragraphs only, no structure inference
- **Combined markdown export**: Exports a combined markdown file (`exported_markdown_content.md`) for content analysis
- **Flexible field mapping**: Use any placeholder names via `scalar_fields` and `block_fields` dictionaries

## Design Philosophy

This project prioritizes:

- Deterministic and reproducible output
- Exact content fidelity (no wording changes)
- Cross-platform Word rendering stability
- Explicit, manual control over list and layout behavior

Design decisions favor predictability and correctness over feature breadth.

## List Rendering Approach

- **Bullet lists**: Manual glyph insertion with configurable glyph sets and indentation policies defined by `ListRenderConfig`. Maximum visual depth is configurable. Deep nesting behavior beyond max_visual_depth is deterministic and policy-driven (clamp_last, cycle, or textual strategies).
- **Numbered lists**: Manual hierarchical numbering (1., 1.1., 1.1.1.) with Python-based counter tracking per list block. Numbers reset for each new list block.
- **Indentation**: Both list types use configurable manual paragraph indentation (default: 0.25" per level, -0.25" hanging indent, configurable via `ListRenderConfig`)
- **No Word numbering XML**: Ensures stable rendering across platforms (especially macOS)

### List Semantics

- **Tree-based rendering** (default): List rendering uses tree-based parsing that attempts to preserve exact order and nesting of list items, allowing for more accurate rendering of complex nested structures.
- **Block-based rendering** (fallback): When `USE_MARKDOWN_TREE_RENDERER=false`, list rendering is block-based, not AST-linked. Each markdown list block is processed independently and rendered as a separate visual block in Word.
- **Visual preservation**: Nested lists are preserved visually (indentation and glyphs reflect nesting depth), not structurally (no Word list object relationships).
- **Mixed lists**: In block-based rendering, mixed bullet ↔ numbered lists are rendered as separate blocks by design. This is a deliberate stability decision to ensure deterministic output and cross-platform consistency. Tree renderer may handle mixed lists differently when enabled.

### Deep Nesting

Very deep nesting (beyond `max_visual_depth`) may reduce readability due to Word layout constraints. This is a limitation of Word document layout, not the export engine. The engine handles deep nesting deterministically according to the configured strategy, but visual clarity may degrade with extreme nesting.

## Compatibility

- **Python**: 3.8+
- **Tested Platforms**: macOS, Linux
- **Windows**: Supported (planned for full testing in future releases)

## Dependencies

- `pydantic>=2.0.0` - For data models
- `python-docx>=1.1.0` - For Word document manipulation
- `markdown-it-py>=3.0.0` - For markdown parsing

## Installation

### From PyPI

```bash
pip install deterministic-docx-export
```

### From Source

```bash
git clone <repository-url>
cd deterministic-docx-export
pip install -r requirements.txt
pip install -e .
```

## Examples

The repository includes example files to help you get started:

- **`example_usage.py`** - Complete usage examples including:
  - Basic usage with dynamic fields
  - Loading data from JSON
  - Plain text mode
  - Observability & export summary (shows how to access summary and log files)
- **`example_data.json`** - Sample JSON input structure

Run examples:
```bash
python example_usage.py
```

## Testing

Run the test suite to verify functionality and inspect outputs:

```bash
# Install test dependencies
pip install -r requirements.txt

# Run all tests
pytest tests/

# Run with verbose output
pytest tests/ -v

# Run a specific test
pytest tests/test_word_export.py::TestBasicExport::test_basic_scalar_fields
```

Test outputs are saved in `tests/test_output/` for inspection:
- Input JSON files (`*_input.json`) - The test data used
- Output DOCX files (`*.docx`) - The generated Word documents
- Markdown files (`*_markdown.md`) - Exported markdown content
- Config files (`*_config.json`) - Configuration used

See `tests/README.md` for more details.

## Usage

### Minimal Example

```python
from pathlib import Path
from deterministic_docx_export.models.export_models import WordExportRequest
from deterministic_docx_export.services.word_export_service import export_to_word

# Create export request
request = WordExportRequest(
    scalar_fields={
        "title": "My Document",
        "author": "John Doe",
    },
    block_fields={
        "content": "# Introduction\n\nThis is the content with **bold** text.",
    }
)

# Export to Word
# Note: If OUTPUT_DIRECTORY environment variable is set, relative paths like
# "output.docx" are resolved relative to that directory. Absolute paths are used as-is.
result = export_to_word(
    template_path=Path("template.docx"),
    request=request,
    markdown_mode=True,
    output_path=Path("output.docx"),
)

print(f"Word file: {result['word_file_path']}")
```

### Basic Example

```python
from pathlib import Path
from deterministic_docx_export.models.export_models import WordExportRequest
from deterministic_docx_export.services.word_export_service import export_to_word

# Create export request with dynamic fields
request = WordExportRequest(
    scalar_fields={
        "document_id": "DOC-12345",
        "title": "My Document",
        "author": "John Doe",
        "date": "2024-01-15",
    },
    block_fields={
        "introduction": "# Introduction\n\nThis is the introduction content...",
        "body": "## Main Content\n\nThis is the main body with **bold** and *italic* text.",
        "conclusion": "## Conclusion\n\nFinal thoughts here...",
    }
)

# Export to Word
# Note: Relative paths are resolved relative to OUTPUT_DIRECTORY (if set) or current directory
# Absolute paths (e.g., Path("/var/exports/doc.docx")) are used as-is
result = export_to_word(
    template_path=Path("template.docx"),
    request=request,
    markdown_mode=True,  # Set to False for plain text mode
    output_path=Path("output/exported_document.docx"),
)

print(f"Word file saved to: {result['word_file_path']}")
if result['markdown_files']:
    print(f"Markdown file saved to: {result['markdown_files'][0]}")
```

### Configuration Example

```python
from pathlib import Path
from deterministic_docx_export.models.export_models import WordExportRequest
from deterministic_docx_export.models.export_config import ListRenderConfig
from deterministic_docx_export.services.word_export_service import export_to_word

# Create custom list rendering configuration
config = ListRenderConfig(
    max_visual_depth=7,  # Allow up to 7 nesting levels
    indent_inches_per_level=0.25,
    hanging_indent_inches=0.25,
    bullet_glyphs=("•", "◦", "▪", "▫", "▸", "▴", "▾"),
    deep_bullet_strategy="clamp_last",  # or "cycle" or "textual"
)

request = WordExportRequest(
    scalar_fields={"title": "Document with Deep Nesting"},
    block_fields={
        "content": """
# Deep Nested Lists

- Level 1
  - Level 2
    - Level 3
      - Level 4
        - Level 5
          - Level 6
            - Level 7
        """
    }
)

result = export_to_word(
    template_path=Path("template.docx"),
    request=request,
    markdown_mode=True,
    output_path=Path("output.docx"),
    config=config,
)
```

### JSON Input Example

```python
import json
from pathlib import Path
from deterministic_docx_export.models.export_models import WordExportRequest
from deterministic_docx_export.services.word_export_service import export_to_word

# Load from JSON
with open("data.json", "r") as f:
    data = json.load(f)

request = WordExportRequest(**data)

result = export_to_word(
    template_path=Path("template.docx"),
    request=request,
    markdown_mode=True,
    output_path=Path("output/exported_document.docx"),
)
```

Example `data.json`:
```json
{
  "scalar_fields": {
    "document_id": "DOC-12345",
    "title": "My Document",
    "author": "John Doe"
  },
  "block_fields": {
    "introduction": "# Introduction\n\nContent here...",
    "body": "## Body\n\nMore content...",
    "conclusion": "## Conclusion\n\nFinal thoughts..."
  }
}
```

## Template Placeholders

### Scalar Placeholders

Scalar placeholders are replaced with simple text values. You can use **any placeholder name** you want by adding it to `scalar_fields`:

```python
scalar_fields={
    "document_id": "DOC-123",
    "title": "My Document",
    "author": "John Doe",
    "custom_field": "Any value",
}
```

In your Word template, use `{{document_id}}`, `{{title}}`, `{{author}}`, `{{custom_field}}`, etc.

### Block Placeholders

Block placeholders are replaced with structured content (markdown or plain text). You can use **any placeholder name** you want by adding it to `block_fields`:

```python
block_fields={
    "introduction": "# Introduction\n\nContent...",
    "body": "## Body\n\nMore content...",
    "conclusion": "## Conclusion\n\nFinal thoughts...",
}
```

In your Word template, use `{{introduction}}`, `{{body}}`, `{{conclusion}}`, etc.

## Markdown Support

When `markdown_mode=True`, the following markdown structures are supported:

- **Headings**: `# H1`, `## H2`, etc. (up to H9)
- **Paragraphs**: Plain text paragraphs
- **Bullet lists**: `-` or `*` for unordered lists
- **Numbered lists**: `1.` for ordered lists
- **Tables**: Markdown table syntax
- **Inline formatting**: `**bold**` and `*italic*` text (in headings and paragraphs, not in table cells)

### Table Behavior

- Inline markdown formatting inside table cells is treated as literal text.
- Tables prioritize structure and layout determinism over inline formatting.

## Output

The export function returns a dictionary with:

- `word_file_path`: Path to the generated DOCX file
- `markdown_files`: List containing the path to `exported_markdown_content.md` (empty if all blocks are empty)
- `export_summary`: (Optional) Export summary with audit trail when `enable_export_trace=True` (default)
- `export_summary_log_path`: (Optional) Path to saved log file when `enable_export_trace=True`

## Observability & Export Summary

The library provides comprehensive observability to track what happened during export execution. This includes:

- **Scalar replacements**: Which placeholders were replaced and where
- **Block replacements**: Which block placeholders were replaced and with what content types
- **Skipped items**: Which placeholders were detected but skipped, and why
- **Fallback events**: When fallback mechanisms were triggered (e.g., tree renderer → block renderer)
- **Warnings**: Safety warnings raised during export

### Quick Start

```python
result = export_to_word(
    template_path=Path("template.docx"),
    request=request,
    markdown_mode=True,
    output_path=Path("output.docx")
)

# Access export summary
if "export_summary" in result:
    summary = result["export_summary"]
    print(f"Scalar replacements: {len(summary['scalar_replacements'])}")
    print(f"Block replacements: {len(summary['block_replacements'])}")
    print(f"Skipped items: {len(summary['skipped_items'])}")
    
    # Log file path
    if "export_summary_log_path" in result:
        print(f"Log saved to: {result['export_summary_log_path']}")
```

### Log Directory Structure

Export summaries are automatically saved to timestamped directories:

```
{EXPORT_LOG_DIR}/logs/
  20240124_143022/
    export_summary_output.json
    export_trace_output.log
```

- Each export run gets its own timestamped directory
- Multiple exports in the same second share the same directory
- Logs are organized chronologically for easy navigation
- Both JSON summary and structured log file are saved automatically

### Environment Variables

Three environment variables control export behavior:

**`OUTPUT_DIRECTORY`**: Base directory for output files (new in v0.3.0)

- Default: Not set (output paths resolved relative to current working directory)
- If set: Relative `output_path` parameters are resolved relative to this directory
- If `output_path` is absolute, this variable is ignored (absolute paths are used as-is)
- Useful for centralizing all exported documents in a single location
- Example: `OUTPUT_DIRECTORY=/var/exports` → `output_path=Path("doc.docx")` becomes `/var/exports/doc.docx`

**`USE_MARKDOWN_TREE_RENDERER`**: Control tree-based rendering engine (enabled by default)

- Default: `True` (tree renderer is enabled by default in v0.3.0)
- Set to `"false"`, `"0"`, or `"no"` to disable and use block-based rendering (legacy behavior)
- When enabled (default), block placeholders are processed using the tree-based renderer
- Tree renderer provides enhanced structure preservation for complex nested markdown
- Falls back to block-based renderer on parsing failures

**`EXPORT_LOG_DIR`**: Controls where export summary logs are saved

- If set: Logs saved to `{EXPORT_LOG_DIR}/logs/{timestamp}/`
- If not set: Logs saved to `{output_dir}/logs/{timestamp}/`

The library automatically loads `.env` files if `python-dotenv` is available.

#### Example `.env` Configuration

```bash
# Base directory for output files (relative paths resolved relative to this)
OUTPUT_DIRECTORY=/var/exports

# Tree renderer is enabled by default, but can be explicitly set:
USE_MARKDOWN_TREE_RENDERER=true

# Or disable to use block-based rendering (legacy behavior):
# USE_MARKDOWN_TREE_RENDERER=false

# Customize log directory
EXPORT_LOG_DIR=/path/to/logs
```

### Enable/Disable

Observability is **enabled by default** (`enable_export_trace=True`). To disable:

```python
request = WordExportRequest(
    scalar_fields={...},
    block_fields={...},
    enable_export_trace=False  # Disable observability
)
```

When disabled, no summary is collected, no structured logs are emitted, and no log files are created (near-zero overhead).

### Documentation

See [`OBSERVABILITY.md`](OBSERVABILITY.md) for complete documentation including:
- Detailed examples
- Summary structure reference
- Log configuration
- API integration
- Testing guidance

## Content Fidelity

This module respects content fidelity: all text is preserved exactly as provided, with only structural transformations applied (markdown syntax → Word objects).

## Design Guarantees & Non-Goals

### Guarantees

- **Deterministic output**: Given the same input, the export produces identical Word documents across runs and platforms.
- **No Word numbering XML**: Lists are rendered using manual glyph/number insertion, avoiding Word's numbering system for cross-platform stability.
- **Cross-platform stability**: Output renders consistently across Windows, macOS, and Linux Word viewers.
- **Exact text fidelity**: All text content is preserved exactly as provided, with no wording modifications.

### Non-Goals

- **Semantic AST reconstruction**: The engine does not attempt to reconstruct semantic markdown AST relationships in Word. Lists are rendered visually, not structurally linked.
- **Word auto-list compatibility**: Lists are not converted to Word's native list objects. This is intentional for stability and determinism.
- **Rich markdown inside tables**: Inline markdown formatting (bold, italic) inside table cells is not processed. Table cells contain literal text only.

## Versioning Strategy

This project follows [Semantic Versioning](https://semver.org/) (MAJOR.MINOR.PATCH):

- **PATCH** (0.1.x): Bug fixes, documentation improvements, internal optimizations
- **MINOR** (0.x.0): New features, enhancements, backward-compatible API additions
- **MAJOR** (x.0.0): Breaking changes, API modifications, changes to deterministic behavior

**Note:** v0.x indicates the API is still evolving. However, **deterministic behavior guarantees will NOT change without a major version bump**.

### Stability Contract

- Deterministic output behavior is guaranteed within the same major version
- API changes within the same major version maintain backward compatibility
- Breaking changes require a major version increment

## License

Licensed under the Apache License, Version 2.0.
You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.

See `LICENSE` and `NOTICE` files for full license text and copyright information.

## Publishing

For maintainers: To publish a new version to PyPI:

```bash
# Clean previous builds
rm -rf build dist *.egg-info

# Build distribution packages
python -m build

# Upload to PyPI (requires PyPI API token)
python -m twine upload dist/*
```

**Note:** Use PyPI API tokens for authentication. See [PyPI documentation](https://pypi.org/help/#apitoken) for token setup.
