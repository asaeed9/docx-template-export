# Deterministic DOCX Export

A standalone Python library for exporting structured content (markdown or plain text) into Word (.docx) templates. This module is **strictly non-generative** - it does not call any LLMs and does not modify wording. It only maps existing content into Word document structures.

## Features

- **Scalar placeholder replacement**: Replace simple placeholders like `{{document_id}}`, `{{title}}`, etc.
- **Block placeholder replacement**: Replace structured content blocks like `{{summary}}`, `{{proposal}}`, etc.
- **Markdown parsing**: Convert markdown content to Word structures (headings, lists, tables)
- **Manual list rendering**: Bullet and numbered lists use deterministic glyph/number insertion (no Word numbering XML) for stable, cross-platform rendering
- **Plain text mode fallback**: Paragraphs only, no structure inference
- **Combined markdown export**: Exports a combined markdown file (`exported_markdown_content.md`) for content analysis
- **Flexible field mapping**: Use any placeholder names via `scalar_fields` and `block_fields` dictionaries

## List Rendering Approach

- **Bullet lists**: Manual glyph insertion with configurable glyph sets and indentation policies defined by `ListRenderConfig`. Maximum visual depth is configurable. Deep nesting behavior beyond max_visual_depth is deterministic and policy-driven (clamp_last, cycle, or textual strategies).
- **Numbered lists**: Manual hierarchical numbering (1., 1.1., 1.1.1.) with Python-based counter tracking per list block. Numbers reset for each new list block.
- **Indentation**: Both list types use configurable manual paragraph indentation (default: 0.25" per level, -0.25" hanging indent, configurable via `ListRenderConfig`)
- **No Word numbering XML**: Ensures stable rendering across platforms (especially macOS)

### List Semantics

- **Block-based rendering**: List rendering is block-based, not AST-linked. Each markdown list block is processed independently and rendered as a separate visual block in Word.
- **Visual preservation**: Nested lists are preserved visually (indentation and glyphs reflect nesting depth), not structurally (no Word list object relationships).
- **Mixed lists**: Mixed bullet ↔ numbered lists are rendered as separate blocks by design. This is a deliberate stability decision to ensure deterministic output and cross-platform consistency. This is not a bug.

### Deep Nesting

Very deep nesting (beyond `max_visual_depth`) may reduce readability due to Word layout constraints. This is a limitation of Word document layout, not the export engine. The engine handles deep nesting deterministically according to the configured strategy, but visual clarity may degrade with extreme nesting.

## Dependencies

- `pydantic` - For data models
- `python-docx` - For Word document manipulation
- `markdown-it-py` - For markdown parsing

## Installation

### From Source

```bash
git clone <repository-url>
cd deterministic-docx-export
pip install -r requirements.txt
```

### As Package (when published)

```bash
pip install deterministic-docx-export
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

### Basic Example

```python
from pathlib import Path
from models.export_models import WordExportRequest
from services.word_export_service import export_to_word

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

### JSON Input Example

```python
import json
from pathlib import Path
from models.export_models import WordExportRequest
from services.word_export_service import export_to_word

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

## License

[Add your license here]
