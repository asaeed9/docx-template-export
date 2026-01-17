"""
Example usage of the deterministic DOCX export service.

This script demonstrates how to use the word export functionality with
dynamic fields for flexible placeholder replacement.
"""
import sys
import json
from pathlib import Path

# Add the parent directory to the path so we can import the modules
sys.path.insert(0, str(Path(__file__).parent))

from models.export_models import WordExportRequest
from services.word_export_service import export_to_word


def example_basic():
    """Basic example using dynamic fields."""
    print("=== Example 1: Basic Usage ===\n")
    
    request = WordExportRequest(
        scalar_fields={
            "document_id": "DOC-12345",
            "title": "Example Document",
            "author": "John Doe",
            "date": "2024-01-15",
            "version": "1.0",
        },
        block_fields={
            "introduction": """
# Introduction

This is the **introduction** content with *formatting* support.

## Key Points

- First important point
- Second important point
  - Nested point
- Third important point
            """,
            "body": """
## Main Content

1. First item
2. Second item
   1. Nested numbered item
   2. Another nested item
3. Third item

### Table Example

| Column 1 | Column 2 |
|----------|----------|
| Data 1   | Data 2   |
| Data 3   | Data 4   |
            """,
            "conclusion": """
## Conclusion

- Summary point 1
- Summary point 2
- Summary point 3
            """,
        },
    )

    template_path = Path("template.docx")
    output_path = Path("output/example_basic.docx")

    try:
        result = export_to_word(
            template_path=template_path,
            request=request,
            markdown_mode=True,
            output_path=output_path,
        )
        print(f"✓ Word file saved to: {result['word_file_path']}")
        if result['markdown_files']:
            print(f"✓ Markdown file saved to: {result['markdown_files'][0]}")
    except FileNotFoundError:
        print(f"⚠ Template file not found at: {template_path}")
        print("   Please create a template.docx file with placeholders like {{title}}, {{introduction}}, etc.")
    except Exception as e:
        print(f"✗ Error: {e}")


def example_from_json():
    """Example loading data from JSON file."""
    print("\n=== Example 2: Loading from JSON ===\n")
    
    json_file = Path("example_data.json")
    if not json_file.exists():
        print(f"⚠ JSON file not found at: {json_file}")
        print("   Skipping this example.")
        return
    
    try:
        with open(json_file, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        request = WordExportRequest(**data)
        
        template_path = Path("template.docx")
        output_path = Path("output/example_from_json.docx")
        
        result = export_to_word(
            template_path=template_path,
            request=request,
            markdown_mode=True,
            output_path=output_path,
        )
        print(f"✓ Word file saved to: {result['word_file_path']}")
        if result['markdown_files']:
            print(f"✓ Markdown file saved to: {result['markdown_files'][0]}")
    except FileNotFoundError as e:
        print(f"⚠ Template file not found: {e}")
    except Exception as e:
        print(f"✗ Error: {e}")


def example_plain_text_mode():
    """Example using plain text mode (no markdown parsing)."""
    print("\n=== Example 3: Plain Text Mode ===\n")
    
    request = WordExportRequest(
        scalar_fields={
            "title": "Plain Text Document",
            "author": "Jane Smith",
        },
        block_fields={
            "content": """
This is plain text content.
It will be split into paragraphs
based on blank lines.

No markdown formatting will be applied.
No lists, no tables, no headings.
Just plain paragraphs.
            """,
        },
    )

    template_path = Path("template.docx")
    output_path = Path("output/example_plain_text.docx")

    try:
        result = export_to_word(
            template_path=template_path,
            request=request,
            markdown_mode=False,  # Plain text mode
            output_path=output_path,
        )
        print(f"✓ Word file saved to: {result['word_file_path']}")
        if result['markdown_files']:
            print(f"✓ Markdown file saved to: {result['markdown_files'][0]}")
    except FileNotFoundError:
        print(f"⚠ Template file not found at: {template_path}")
    except Exception as e:
        print(f"✗ Error: {e}")


def main():
    """Run all examples."""
    print("Deterministic DOCX Export - Usage Examples\n")
    print("=" * 60 + "\n")
    
    example_basic()
    example_from_json()
    example_plain_text_mode()
    
    print("\n" + "=" * 60)
    print("\nNote: Make sure you have a template.docx file with appropriate")
    print("placeholders (e.g., {{title}}, {{introduction}}, etc.) in your")
    print("current directory or update the template_path in the examples.")


if __name__ == "__main__":
    main()
