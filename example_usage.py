"""
Example usage of the deterministic DOCX export service.

This script demonstrates how to use the word export functionality with
dynamic fields for flexible placeholder replacement.
"""
import json
from pathlib import Path

from docx_template_export.models.export_models import WordExportRequest
from docx_template_export.services.word_export_service import export_to_word


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


def example_observability():
    """Example demonstrating observability features: export summary and log files."""
    print("\n=== Example 4: Observability & Export Summary ===\n")
    
    import logging
    
    # Configure logging to see structured logs in console (optional)
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    
    request = WordExportRequest(
        scalar_fields={
            "document_id": "DOC-12345",
            "title": "Observability Example",
            "author": "John Doe"
        },
        block_fields={
            "summary": """
# Summary

This is a **summary** with *formatting*.

- Point 1
- Point 2
- Point 3
            """,
            "details": "## Details\n\nMore content here..."
        },
        enable_export_trace=True  # This is the default - enables summary collection and log files
    )

    template_path = Path("template.docx")
    output_path = Path("output/example_observability.docx")

    try:
        result = export_to_word(
            template_path=template_path,
            request=request,
            markdown_mode=True,
            output_path=output_path,
        )
        
        print(f"✓ Word file saved to: {result['word_file_path']}")
        if result.get('markdown_files'):
            print(f"✓ Markdown file saved to: {result['markdown_files'][0]}")
        
        # Access export summary (automatically collected when enable_export_trace=True)
        if "export_summary" in result:
            summary = result["export_summary"]
            
            print("\n" + "=" * 60)
            print("📊 EXPORT SUMMARY (AUDIT TRAIL)")
            print("=" * 60)
            
            print(f"\n📝 Scalar Replacements: {len(summary['scalar_replacements'])}")
            for i, replacement in enumerate(summary['scalar_replacements'], 1):
                print(f"   {i}. {replacement['token']} → {replacement['value_preview']}")
                print(f"      Location: {replacement['region']}/{replacement['container']}")
            
            print(f"\n📦 Block Replacements: {len(summary['block_replacements'])}")
            for i, replacement in enumerate(summary['block_replacements'], 1):
                print(f"   {i}. {replacement['token']}")
                print(f"      Location: {replacement['region']}/{replacement['container']}")
                print(f"      Block types: {', '.join(replacement['block_types'])}")
            
            print(f"\n⏭️  Skipped Items: {len(summary['skipped_items'])}")
            for i, skipped in enumerate(summary['skipped_items'], 1):
                print(f"   {i}. {skipped['token']} ({skipped['reason']})")
            
            print(f"\n🔄 Fallback Events: {len(summary['fallback_events'])}")
            for i, event in enumerate(summary['fallback_events'], 1):
                print(f"   {i}. {event['event_type']}: {event.get('reason', 'N/A')}")
            
            print(f"\n⚠️  Warnings: {len(summary['warnings'])}")
            for i, warning in enumerate(summary['warnings'], 1):
                print(f"   {i}. {warning['message']}")
            
            # Show log file locations
            print("\n" + "=" * 60)
            print("📁 LOG FILES (Automatically Saved)")
            print("=" * 60)
            
            if "export_summary_log_path" in result:
                json_log_path = Path(result["export_summary_log_path"])
                print(f"\n✓ JSON Summary: {json_log_path}")
                
                # The structured log file is in the same directory
                log_dir = json_log_path.parent
                output_filename = output_path.stem
                trace_log_path = log_dir / f"export_trace_{output_filename}.log"
                
                if trace_log_path.exists():
                    print(f"✓ Structured Log: {trace_log_path}")
                    print(f"\n  Both files are in: {log_dir}")
                    print(f"  - JSON summary: Complete audit trail in structured format")
                    print(f"  - Log file: All INFO-level messages from the export process")
                else:
                    print(f"⚠ Structured log file not found (may not have been created)")
            
            # Optionally save summary to a custom location
            custom_summary_file = Path("export_audit_custom.json")
            with open(custom_summary_file, "w") as f:
                json.dump(summary, f, indent=2)
            print(f"\n💾 Summary also saved to: {custom_summary_file}")
            
        else:
            print("\n⚠️  No export summary found (enable_export_trace=False)")
            print("   Set enable_export_trace=True to enable observability features")
            
    except FileNotFoundError:
        print(f"⚠ Template file not found at: {template_path}")
        print("\nNote: Even without a template, this example shows:")
        print("  - How to access result['export_summary']")
        print("  - Where log files are automatically saved")
        print("  - How to view the audit trail")
    except Exception as e:
        print(f"✗ Error: {e}")
        import traceback
        traceback.print_exc()


def main():
    """Run all examples."""
    print("Deterministic DOCX Export - Usage Examples\n")
    print("=" * 60 + "\n")
    
    example_basic()
    example_from_json()
    example_plain_text_mode()
    example_observability()
    
    print("\n" + "=" * 60)
    print("\nNote: Make sure you have a template.docx file with appropriate")
    print("placeholders (e.g., {{title}}, {{introduction}}, etc.) in your")
    print("current directory or update the template_path in the examples.")
    print("\nObservability: Export summaries and log files are automatically")
    print("saved to logs/{timestamp}/ when enable_export_trace=True (default).")


if __name__ == "__main__":
    main()
