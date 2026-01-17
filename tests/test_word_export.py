"""
Tests for word export functionality.

Test outputs are saved in tests/test_output/ for inspection.
"""
import json
from pathlib import Path

import pytest

from models.export_models import WordExportRequest
from models.export_config import ListRenderConfig
from services.word_export_service import export_to_word


class TestBasicExport:
    """Test basic export functionality."""
    
    def test_basic_scalar_fields(self, simple_template, test_output_dir):
        """Test basic scalar field replacement."""
        request = WordExportRequest(
            scalar_fields={
                "document_id": "TEST-001",
                "title": "Test Document",
                "author": "Test Author",
                "date": "2024-01-15",
            },
            block_fields={},
        )
        
        output_path = test_output_dir / "test_basic_scalar.docx"
        
        result = export_to_word(
            template_path=simple_template,
            request=request,
            markdown_mode=False,
            output_path=output_path,
        )
        
        assert result["word_file_path"] == str(output_path)
        assert output_path.exists()
        
        # Save test input for inspection
        input_file = test_output_dir / "test_basic_scalar_input.json"
        with open(input_file, "w", encoding="utf-8") as f:
            json.dump(request.model_dump(), f, indent=2, ensure_ascii=False)
    
    def test_basic_block_fields(self, simple_template, test_output_dir):
        """Test basic block field replacement."""
        request = WordExportRequest(
            scalar_fields={
                "document_id": "TEST-002",
                "title": "Block Fields Test",
                "author": "Test Author",
                "date": "2024-01-15",
            },
            block_fields={
                "introduction": "# Introduction\n\nThis is a test introduction.",
                "body": "## Body\n\nThis is the body content.",
                "conclusion": "## Conclusion\n\nThis is the conclusion.",
            },
        )
        
        output_path = test_output_dir / "test_basic_blocks.docx"
        
        result = export_to_word(
            template_path=simple_template,
            request=request,
            markdown_mode=True,
            output_path=output_path,
        )
        
        assert result["word_file_path"] == str(output_path)
        assert output_path.exists()
        assert len(result["markdown_files"]) == 1
        
        # Save test input for inspection
        input_file = test_output_dir / "test_basic_blocks_input.json"
        with open(input_file, "w", encoding="utf-8") as f:
            json.dump(request.model_dump(), f, indent=2, ensure_ascii=False)


class TestMarkdownRendering:
    """Test markdown rendering features."""
    
    def test_headings(self, simple_template, test_output_dir):
        """Test heading rendering."""
        request = WordExportRequest(
            scalar_fields={"document_id": "TEST-003", "title": "Headings Test"},
            block_fields={
                "introduction": """
# Heading 1

## Heading 2

### Heading 3

#### Heading 4
                """,
            },
        )
        
        output_path = test_output_dir / "test_headings.docx"
        
        result = export_to_word(
            template_path=simple_template,
            request=request,
            markdown_mode=True,
            output_path=output_path,
        )
        
        assert output_path.exists()
        
        # Save test input
        input_file = test_output_dir / "test_headings_input.json"
        with open(input_file, "w", encoding="utf-8") as f:
            json.dump(request.model_dump(), f, indent=2, ensure_ascii=False)
    
    def test_bullet_lists(self, simple_template, test_output_dir):
        """Test bullet list rendering."""
        request = WordExportRequest(
            scalar_fields={"document_id": "TEST-004", "title": "Bullet Lists Test"},
            block_fields={
                "introduction": """
# Bullet Lists

- First item
- Second item
  - Nested item 1
  - Nested item 2
- Third item
  - Deep nested item
                """,
            },
        )
        
        output_path = test_output_dir / "test_bullet_lists.docx"
        
        result = export_to_word(
            template_path=simple_template,
            request=request,
            markdown_mode=True,
            output_path=output_path,
        )
        
        assert output_path.exists()
        
        # Save test input
        input_file = test_output_dir / "test_bullet_lists_input.json"
        with open(input_file, "w", encoding="utf-8") as f:
            json.dump(request.model_dump(), f, indent=2, ensure_ascii=False)
    
    def test_numbered_lists(self, simple_template, test_output_dir):
        """Test numbered list rendering."""
        request = WordExportRequest(
            scalar_fields={"document_id": "TEST-005", "title": "Numbered Lists Test"},
            block_fields={
                "introduction": """
# Numbered Lists

1. First item
2. Second item
   1. Nested item 1
   2. Nested item 2
3. Third item
                """,
            },
        )
        
        output_path = test_output_dir / "test_numbered_lists.docx"
        
        result = export_to_word(
            template_path=simple_template,
            request=request,
            markdown_mode=True,
            output_path=output_path,
        )
        
        assert output_path.exists()
        
        # Save test input
        input_file = test_output_dir / "test_numbered_lists_input.json"
        with open(input_file, "w", encoding="utf-8") as f:
            json.dump(request.model_dump(), f, indent=2, ensure_ascii=False)
    
    def test_tables(self, simple_template, test_output_dir):
        """Test table rendering."""
        request = WordExportRequest(
            scalar_fields={"document_id": "TEST-006", "title": "Tables Test"},
            block_fields={
                "introduction": """
# Tables

| Column 1 | Column 2 | Column 3 |
|----------|----------|----------|
| Row 1 Col 1 | Row 1 Col 2 | Row 1 Col 3 |
| Row 2 Col 1 | Row 2 Col 2 | Row 2 Col 3 |
| Row 3 Col 1 | Row 3 Col 2 | Row 3 Col 3 |
                """,
            },
        )
        
        output_path = test_output_dir / "test_tables.docx"
        
        result = export_to_word(
            template_path=simple_template,
            request=request,
            markdown_mode=True,
            output_path=output_path,
        )
        
        assert output_path.exists()
        
        # Save test input
        input_file = test_output_dir / "test_tables_input.json"
        with open(input_file, "w", encoding="utf-8") as f:
            json.dump(request.model_dump(), f, indent=2, ensure_ascii=False)
    
    def test_formatting(self, simple_template, test_output_dir):
        """Test text formatting (bold, italic)."""
        request = WordExportRequest(
            scalar_fields={"document_id": "TEST-007", "title": "Formatting Test"},
            block_fields={
                "introduction": """
# Formatting

This is **bold text** and this is *italic text*.

This is ***bold and italic*** text.

Normal text with **bold** in the middle and *italic* too.
                """,
            },
        )
        
        output_path = test_output_dir / "test_formatting.docx"
        
        result = export_to_word(
            template_path=simple_template,
            request=request,
            markdown_mode=True,
            output_path=output_path,
        )
        
        assert output_path.exists()
        
        # Save test input
        input_file = test_output_dir / "test_formatting_input.json"
        with open(input_file, "w", encoding="utf-8") as f:
            json.dump(request.model_dump(), f, indent=2, ensure_ascii=False)


class TestPlainTextMode:
    """Test plain text mode (no markdown parsing)."""
    
    def test_plain_text_mode(self, simple_template, test_output_dir):
        """Test plain text mode rendering."""
        request = WordExportRequest(
            scalar_fields={"document_id": "TEST-008", "title": "Plain Text Test"},
            block_fields={
                "introduction": """
This is plain text.
It will be split into paragraphs
based on blank lines.

No markdown formatting will be applied.
                """,
            },
        )
        
        output_path = test_output_dir / "test_plain_text.docx"
        
        result = export_to_word(
            template_path=simple_template,
            request=request,
            markdown_mode=False,
            output_path=output_path,
        )
        
        assert output_path.exists()
        
        # Save test input
        input_file = test_output_dir / "test_plain_text_input.json"
        with open(input_file, "w", encoding="utf-8") as f:
            json.dump(request.model_dump(), f, indent=2, ensure_ascii=False)


class TestListRenderConfig:
    """Test list rendering configuration."""
    
    def test_custom_list_config(self, simple_template, test_output_dir):
        """Test custom list rendering configuration."""
        config = ListRenderConfig(
            max_visual_depth=5,
            indent_inches_per_level=0.5,
            hanging_indent_inches=0.3,
            bullet_glyphs=("•", "◦", "▪", "▫", "▸"),
            deep_bullet_strategy="cycle",
        )
        
        request = WordExportRequest(
            scalar_fields={"document_id": "TEST-009", "title": "Custom Config Test"},
            block_fields={
                "introduction": """
# Deep Nesting Test

- Level 1
  - Level 2
    - Level 3
      - Level 4
        - Level 5
          - Level 6 (should cycle)
            - Level 7 (should cycle)
                """,
            },
        )
        
        output_path = test_output_dir / "test_custom_config.docx"
        
        result = export_to_word(
            template_path=simple_template,
            request=request,
            markdown_mode=True,
            output_path=output_path,
            config=config,
        )
        
        assert output_path.exists()
        
        # Save test input and config
        input_file = test_output_dir / "test_custom_config_input.json"
        with open(input_file, "w", encoding="utf-8") as f:
            json.dump(request.model_dump(), f, indent=2, ensure_ascii=False)
        
        config_file = test_output_dir / "test_custom_config_config.json"
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump({
                "max_visual_depth": config.max_visual_depth,
                "indent_inches_per_level": config.indent_inches_per_level,
                "hanging_indent_inches": config.hanging_indent_inches,
                "bullet_glyphs": config.bullet_glyphs,
                "deep_bullet_strategy": config.deep_bullet_strategy,
            }, f, indent=2, ensure_ascii=False)
    
    def test_deep_nesting_with_default_config(self, simple_template, test_output_dir):
        """Test that deep nesting (beyond level 3) is preserved and handled correctly."""
        # Use default config (max_visual_depth=3)
        request = WordExportRequest(
            scalar_fields={"document_id": "TEST-DEEP-001", "title": "Deep Nesting Test (Default Config)"},
            block_fields={
                "introduction": """
# Deep Nesting Test

This test verifies that levels beyond 3 are properly handled:

- Level 1 item
  - Level 2 item
    - Level 3 item
      - Level 4 item (should be clamped to 3 with default config)
        - Level 5 item (should be clamped to 3)
          - Level 6 item (should be clamped to 3)
            - Level 7 item (should be clamped to 3)
                """,
            },
        )
        
        output_path = test_output_dir / "test_deep_nesting_default.docx"
        
        result = export_to_word(
            template_path=simple_template,
            request=request,
            markdown_mode=True,
            output_path=output_path,
        )
        
        assert output_path.exists()
        
        # Save test input
        input_file = test_output_dir / "test_deep_nesting_default_input.json"
        with open(input_file, "w", encoding="utf-8") as f:
            json.dump(request.model_dump(), f, indent=2, ensure_ascii=False)
    
    def test_deep_nesting_with_extended_config(self, simple_template, test_output_dir):
        """Test deep nesting with extended max_visual_depth."""
        config = ListRenderConfig(
            max_visual_depth=7,  # Allow up to 7 levels
            indent_inches_per_level=0.25,
            hanging_indent_inches=0.25,
            bullet_glyphs=("•", "◦", "▪", "▫", "▸", "▴", "▾"),
            deep_bullet_strategy="clamp_last",
        )
        
        request = WordExportRequest(
            scalar_fields={"document_id": "TEST-DEEP-002", "title": "Deep Nesting Test (Extended Config)"},
            block_fields={
                "introduction": """
# Deep Nesting Test (Extended)

This test uses max_visual_depth=7 to show deeper nesting:

- Level 1
  - Level 2
    - Level 3
      - Level 4
        - Level 5
          - Level 6
            - Level 7
              - Level 8 (should be clamped to 7)
                - Level 9 (should be clamped to 7)
                """,
            },
        )
        
        output_path = test_output_dir / "test_deep_nesting_extended.docx"
        
        result = export_to_word(
            template_path=simple_template,
            request=request,
            markdown_mode=True,
            output_path=output_path,
            config=config,
        )
        
        assert output_path.exists()
        
        # Save test input and config
        input_file = test_output_dir / "test_deep_nesting_extended_input.json"
        with open(input_file, "w", encoding="utf-8") as f:
            json.dump(request.model_dump(), f, indent=2, ensure_ascii=False)
        
        config_file = test_output_dir / "test_deep_nesting_extended_config.json"
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump({
                "max_visual_depth": config.max_visual_depth,
                "indent_inches_per_level": config.indent_inches_per_level,
                "hanging_indent_inches": config.hanging_indent_inches,
                "bullet_glyphs": config.bullet_glyphs,
                "deep_bullet_strategy": config.deep_bullet_strategy,
            }, f, indent=2, ensure_ascii=False)


class TestComplexMarkdown:
    """Test complex markdown with deep nesting and mixed structures."""
    
    def test_complex_nested_markdown(self, comprehensive_template, test_output_dir):
        """Test extremely complex markdown with deep nesting, mixed lists, and long content."""
        request = WordExportRequest(
            scalar_fields={
                "document_id": "COMPLEX-TEST-001",
                "title": "Complex Nested Markdown Test",
                "author": "Test Suite",
                "version": "1.0.0",
            },
            block_fields={
                "introduction": """
# Complex Nested Markdown Test

This document tests **extremely complex** markdown structures with:

- Deep nesting (up to 5+ levels)
- Mixed list types (bullets and numbered in same section)
- Long content sections
- Multiple formatting styles
- Complex table structures
- Nested lists within lists

## Overview

This test is designed to push the limits of the markdown parser and renderer.
                """,
                "body": """
## Section 1: Deep Nested Bullet Lists

### Level 1 Bullet
- First top-level item with **bold** text
  - Second level item with *italic* text
    - Third level item
      - Fourth level item
        - Fifth level item
          - Sixth level item (should be handled by deep strategy)
            - Seventh level item (should cycle or clamp)
              - Eighth level item (extreme nesting)
  - Another second level item
    - Nested third level
      - Nested fourth level
- Back to first level
  - New second level branch
    - Third level in new branch
      - Fourth level
        - Fifth level
          - Sixth level

### Level 1 Numbered
1. First numbered item
   1. Nested numbered item 1.1
      1. Deep nested 1.1.1
         1. Very deep 1.1.1.1
            1. Extremely deep 1.1.1.1.1
               1. Maximum depth 1.1.1.1.1.1
   2. Nested numbered item 1.2
      1. Deep nested 1.2.1
      2. Deep nested 1.2.2
2. Second numbered item
   1. Nested 2.1
   2. Nested 2.2
      1. Deep 2.2.1
      2. Deep 2.2.2
3. Third numbered item

## Section 2: Mixed Lists in Single Section

This section contains **both bullet and numbered lists** mixed together:

### Mixed Content Block

1. First numbered point
   - Bullet sub-point A
   - Bullet sub-point B
     1. Numbered sub-sub-point B.1
     2. Numbered sub-sub-point B.2
        - Bullet sub-sub-sub-point
          - Deep bullet point
            - Very deep bullet
   - Bullet sub-point C
2. Second numbered point
   - Bullet A
   - Bullet B
     - Nested bullet B.1
     - Nested bullet B.2
       1. Numbered within bullet
       2. Another numbered item
          - Bullet within numbered within bullet
3. Third numbered point

### Another Mixed Section

- Bullet item 1
  1. Numbered under bullet 1.1
  2. Numbered under bullet 1.2
     - Bullet under numbered 1.2.1
     - Bullet under numbered 1.2.2
       1. Numbered under bullet under numbered
       2. Another numbered
          - Bullet under numbered under bullet under numbered
- Bullet item 2
  1. Numbered 2.1
  2. Numbered 2.2
- Bullet item 3

## Section 3: Complex Tables

### Table with Multiple Rows and Columns

| Feature | Status | Notes | Priority | Owner |
|---------|--------|-------|----------|-------|
| Deep Nesting | ✓ | Supports up to configurable depth | High | Team A |
| Mixed Lists | ✓ | Bullets and numbered can be mixed | High | Team A |
| Formatting | ✓ | Bold and italic supported | Medium | Team B |
| Tables | ✓ | Full markdown table support | Medium | Team B |
| Headings | ✓ | H1 through H9 supported | Low | Team C |
| Long Content | ✓ | Handles long paragraphs | Low | Team C |

### Nested Table Information

| Category | Sub-category | Details | Count |
|----------|--------------|---------|-------|
| Lists | Bullet | Simple bullets | 50+ |
| Lists | Bullet | Nested bullets | 30+ |
| Lists | Numbered | Simple numbered | 40+ |
| Lists | Numbered | Nested numbered | 25+ |
| Lists | Mixed | Bullet + Numbered | 15+ |
| Formatting | Bold | **Bold text** | 20+ |
| Formatting | Italic | *Italic text* | 20+ |
| Formatting | Both | ***Bold and italic*** | 10+ |

## Section 4: Long Content with Formatting

### Paragraph with Extensive Formatting

This is a **very long paragraph** that contains *multiple formatting styles* throughout. 
We have **bold text** here and *italic text* there, and even some ***bold and italic combined*** text. 
The paragraph continues with more content to test how the system handles **long-form text** 
with *intermittent formatting* that spans multiple sentences and includes various 
**formatting combinations** to ensure proper rendering.

### Another Long Section

This section demonstrates **complex nested structures** within longer content blocks. 
We have:

1. First major point with **bold emphasis**
   - Sub-point A with *italic text*
   - Sub-point B with **bold text**
     1. Sub-sub-point with ***bold and italic***
     2. Another sub-sub-point
        - Deep bullet point
          - Very deep bullet
            - Extremely deep bullet
   - Sub-point C
2. Second major point
   - Sub-point D
   - Sub-point E
     1. Numbered item E.1
     2. Numbered item E.2
        - Bullet under E.2
          - Deep bullet
3. Third major point

The content continues here with more **formatted text** and *various styles* to test 
the system's ability to handle complex, mixed content structures.

## Section 5: Extreme Nesting Test

### Maximum Depth Testing

- Level 1
  - Level 2
    - Level 3
      - Level 4
        - Level 5
          - Level 6
            - Level 7
              - Level 8
                - Level 9
                  - Level 10 (extreme)

1. Numbered Level 1
   1. Numbered Level 2
      1. Numbered Level 3
         1. Numbered Level 4
            1. Numbered Level 5
               1. Numbered Level 6
                  1. Numbered Level 7
                     1. Numbered Level 8
                        1. Numbered Level 9
                           1. Numbered Level 10 (extreme)

## Section 6: Formatting Combinations

### All Formatting Styles

This section tests **all possible formatting combinations**:

- **Bold bullet item**
- *Italic bullet item*
- ***Bold and italic bullet item***
- Normal bullet item with **bold** in the middle
- Normal bullet item with *italic* in the middle
- Normal bullet item with ***bold and italic*** in the middle

1. **Bold numbered item**
2. *Italic numbered item*
3. ***Bold and italic numbered item***
4. Normal numbered item with **bold** in the middle
5. Normal numbered item with *italic* in the middle
6. Normal numbered item with ***bold and italic*** in the middle

## Section 7: Real-World Scenario

### Project Requirements Document

This simulates a **real-world use case** with complex nested requirements:

#### Functional Requirements

1. User Authentication
   - Login functionality
     1. Email/password login
     2. OAuth integration
        - Google OAuth
        - GitHub OAuth
          - Personal accounts
          - Organization accounts
   - Registration process
     1. Email verification
     2. Profile setup
        - Basic information
        - Preferences
   - Password management
     - Reset password
     - Change password

2. Data Management
   - CRUD operations
     1. Create operations
        - Form validation
        - Data sanitization
     2. Read operations
        - Pagination
        - Filtering
          - By date
          - By category
          - By status
     3. Update operations
     4. Delete operations
   - Data export
     - CSV export
     - PDF export
       - With formatting
       - Without formatting

#### Technical Requirements

- Performance
  1. Response time < 200ms
  2. Database optimization
     - Indexing strategy
     - Query optimization
  3. Caching strategy
    - Redis caching
    - CDN for static assets
- Security
  1. Encryption
     - Data at rest
     - Data in transit
  2. Authentication
     - JWT tokens
     - Session management
  3. Authorization
    - Role-based access
    - Permission system

#### Testing Requirements

| Test Type | Coverage | Priority | Status |
|-----------|----------|----------|--------|
| Unit Tests | 80%+ | High | In Progress |
| Integration Tests | 70%+ | High | Planned |
| E2E Tests | 60%+ | Medium | Planned |
| Performance Tests | 100% | High | Complete |
| Security Tests | 90%+ | Critical | In Progress |

### Implementation Details

The implementation follows these **key principles**:

1. **Modularity**
   - Component-based architecture
   - Service layer separation
   - Clear interfaces
2. **Scalability**
   - Horizontal scaling support
   - Database sharding
   - Load balancing
3. **Maintainability**
   - Comprehensive documentation
   - Code reviews
   - Automated testing
                """,
                "conclusion": """
## Conclusion

This **complex test** demonstrates the library's ability to handle:

- ✓ Deep nesting (10+ levels)
- ✓ Mixed list types
- ✓ Long content sections
- ✓ Complex formatting
- ✓ Multiple table structures
- ✓ Real-world scenarios

### Final Summary

1. **Nesting**: Successfully handles extreme nesting depths
2. **Mixed Lists**: Properly renders bullets and numbered lists together
3. **Formatting**: Correctly applies bold, italic, and combined formatting
4. **Tables**: Renders complex multi-column tables
5. **Long Content**: Handles extensive paragraphs and sections

### Test Results

| Aspect | Result | Notes |
|--------|--------|-------|
| Deep Nesting | ✓ | Handles 10+ levels with appropriate strategy |
| Mixed Lists | ✓ | Bullets and numbered lists work together |
| Formatting | ✓ | Bold, italic, and combinations work correctly |
| Tables | ✓ | Complex tables render properly |
| Long Content | ✓ | Extensive content handled without issues |

**This test validates the robustness of the markdown rendering system.**
                """,
            },
        )
        
        output_path = test_output_dir / "test_complex_nested_markdown.docx"
        
        # Use extended config to allow deep nesting (up to 10 levels)
        config = ListRenderConfig(
            max_visual_depth=10,  # Allow up to 10 levels for complex test
            indent_inches_per_level=0.25,
            hanging_indent_inches=0.25,
            bullet_glyphs=("•", "◦", "▪", "▫", "▸", "▴", "▾", "▶", "◀", "◆"),
            deep_bullet_strategy="clamp_last",
        )
        
        result = export_to_word(
            template_path=comprehensive_template,
            request=request,
            markdown_mode=True,
            output_path=output_path,
            config=config,
        )
        
        assert output_path.exists()
        assert len(result["markdown_files"]) == 1
        
        # Save test input
        input_file = test_output_dir / "test_complex_nested_markdown_input.json"
        with open(input_file, "w", encoding="utf-8") as f:
            json.dump(request.model_dump(), f, indent=2, ensure_ascii=False)
        
        # Save config
        config_file = test_output_dir / "test_complex_nested_markdown_config.json"
        with open(config_file, "w", encoding="utf-8") as f:
            json.dump({
                "max_visual_depth": config.max_visual_depth,
                "indent_inches_per_level": config.indent_inches_per_level,
                "hanging_indent_inches": config.hanging_indent_inches,
                "bullet_glyphs": config.bullet_glyphs,
                "deep_bullet_strategy": config.deep_bullet_strategy,
            }, f, indent=2, ensure_ascii=False)
        
        # Save readable markdown
        if result["markdown_files"]:
            markdown_path = Path(result["markdown_files"][0])
            if markdown_path.exists():
                readable_md = test_output_dir / "test_complex_nested_markdown.md"
                readable_md.write_text(markdown_path.read_text(encoding="utf-8"), encoding="utf-8")
        
        print(f"\n✓ Complex test generated: {output_path}")
        print(f"  Input saved to: {input_file}")
        print(f"  Config saved to: {config_file}")
        if result["markdown_files"]:
            print(f"  Markdown saved to: {test_output_dir / 'test_complex_nested_markdown.md'}")


class TestComprehensive:
    """Comprehensive integration tests."""
    
    def test_comprehensive_export(self, comprehensive_template, test_output_dir):
        """Test comprehensive export with all features."""
        request = WordExportRequest(
            scalar_fields={
                "document_id": "TEST-COMPREHENSIVE-001",
                "title": "Comprehensive Test Document",
                "author": "Test Suite",
                "version": "1.0.0",
            },
            block_fields={
                "introduction": """
# Introduction

This is a **comprehensive** test that includes:

- Multiple features
- Various markdown elements
- Complex structures

## Overview

1. First point
2. Second point
3. Third point
                """,
                "body": """
## Main Content

This section contains the main body of the document.

### Details

| Feature | Status | Notes |
|---------|--------|-------|
| Headings | ✓ | Working |
| Lists | ✓ | Working |
| Tables | ✓ | Working |
| Formatting | ✓ | Working |

### Additional Information

- Item A
- Item B
  - Sub-item B1
  - Sub-item B2
                """,
                "conclusion": """
## Conclusion

This is the conclusion section with *italic* and **bold** text.

### Summary

1. Feature 1: Complete
2. Feature 2: Complete
3. Feature 3: Complete
                """,
                "custom_section": """
# Custom Section

This is a custom section that demonstrates flexibility.
                """,
            },
        )
        
        output_path = test_output_dir / "test_comprehensive.docx"
        
        result = export_to_word(
            template_path=comprehensive_template,
            request=request,
            markdown_mode=True,
            output_path=output_path,
        )
        
        assert output_path.exists()
        assert len(result["markdown_files"]) == 1
        
        # Save test input
        input_file = test_output_dir / "test_comprehensive_input.json"
        with open(input_file, "w", encoding="utf-8") as f:
            json.dump(request.model_dump(), f, indent=2, ensure_ascii=False)
        
        # Also save a readable version of the markdown output
        if result["markdown_files"]:
            markdown_path = Path(result["markdown_files"][0])
            if markdown_path.exists():
                # Copy to test output with a readable name
                readable_md = test_output_dir / "test_comprehensive_markdown.md"
                readable_md.write_text(markdown_path.read_text(encoding="utf-8"), encoding="utf-8")
