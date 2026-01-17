# Test Suite

This directory contains tests for the deterministic DOCX export library.

## Running Tests

```bash
# Run all tests
pytest tests/

# Run with verbose output
pytest tests/ -v

# Run a specific test
pytest tests/test_word_export.py::TestBasicExport::test_basic_scalar_fields

# Run with output display
pytest tests/ -v -s
```

## Test Output

All test outputs are saved in `tests/test_output/` for inspection:

- **Input files**: `*_input.json` - The JSON input used for each test
- **Output files**: `*.docx` - The generated Word documents
- **Markdown files**: `*_markdown.md` - The exported markdown content (for some tests)
- **Config files**: `*_config.json` - Configuration used (for config tests)

## Test Structure

- `test_basic_*.docx` - Basic functionality tests
- `test_headings.docx` - Heading rendering
- `test_bullet_lists.docx` - Bullet list rendering
- `test_numbered_lists.docx` - Numbered list rendering
- `test_tables.docx` - Table rendering
- `test_formatting.docx` - Text formatting (bold/italic)
- `test_plain_text.docx` - Plain text mode
- `test_custom_config.docx` - Custom list rendering configuration
- `test_comprehensive.docx` - Comprehensive integration test

## Inspecting Results

After running tests, you can:

1. Open the generated `.docx` files in Microsoft Word or LibreOffice
2. Compare the `*_input.json` files with the output to verify correctness
3. Check the markdown files to see the raw content that was exported
4. Review the config files to understand custom configurations

## Test Fixtures

Templates are automatically generated in `tests/fixtures/`:
- `simple_template.docx` - Basic template with common placeholders
- `comprehensive_template.docx` - Template with multiple sections
