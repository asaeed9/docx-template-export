# Observability
This library includes **opt-in runtime observability** (enabled by default) that records what happened during a DOCX export run—without changing export output.

The goal is to provide an **audit trail** for:
- What placeholders were replaced (and where)
- What placeholders were detected but skipped (and why)
- What fallback paths were taken
- What warnings were raised

The summary is gated by `enable_export_trace` in `WordExportRequest`.

---

## Quick start

```python
from pathlib import Path

from docx_template_export import export_to_word
from docx_template_export.models import WordExportRequest

request = WordExportRequest(
    scalar_fields={"src_no": "2025-014"},
    block_fields={"summary": "# Background\n\nSome text"},
    enable_export_trace=True,  # default is True
)

result = export_to_word(
    template_path=Path("template.docx"),
    request=request,
    markdown_mode=True,
    output_path=Path("output.docx"),
)

# Access export summary (when enabled)
summary = result.get("export_summary")
if summary:
    print(f"Scalar replacements: {len(summary.get('scalar_replacements', []))}")
    print(f"Block replacements: {len(summary.get('block_replacements', []))}")
    print(f"Skipped items: {len(summary.get('skipped_items', []))}")
    print(f"Fallback events: {len(summary.get('fallback_events', []))}")
    print(f"Warnings: {len(summary.get('warnings', []))}")

print("Summary JSON path:", result.get("export_summary_log_path"))
```

---

## What gets recorded

The structured summary is implemented in `docx_template_export/models/export_summary.py` as `ExportSummary`.

### Regions
Each event records **where** it occurred:
- `BODY`: main document body
- `HEADER`: header region (scalar-only)
- `FOOTER`: footer region (scalar-only)

### Containers
Each event records **what kind of container** it occurred in:
- `FLOW`: normal paragraph flow
- `TABLE`: table cell content
- `TEXTBOX`: textbox content (scalar-only)

### Summary schema (JSON)
`ExportSummary.to_dict()` serializes to the following structure:

```json
{
  "scalar_replacements": [
    {"token": "{{title}}", "region": "BODY", "container": "FLOW", "value_preview": "My Title"}
  ],
  "block_replacements": [
    {"token": "{{summary}}", "region": "BODY", "container": "FLOW", "block_types": ["heading", "paragraph", "bullet_list"]}
  ],
  "skipped_items": [
    {"token": "{{summary}}", "region": "HEADER", "container": "FLOW", "reason": "block_in_header"}
  ],
  "fallback_events": [
    {"event_type": "tree_parse_failed", "location": "{{summary}}", "reason": "parse_markdown_to_tree returned None"}
  ],
  "warnings": [
    {"message": "Block placeholder found in textbox, skipping expansion", "context": {"token": "{{summary}}"}}
  ]
}
```

Notes:
- `token` values include braces, e.g. `{{summary}}`.
- `value_preview` is a short audit-safe preview (not full content).

---

## Skips & common reasons

Placeholders may be detected but **intentionally skipped** for safety and determinism. Common skip reasons include:
- `block_in_textbox`: block placeholder found inside a textbox (textboxes are scalar-only)
- `block_in_header`: block placeholder found in header (headers are scalar-only)
- `block_in_footer`: block placeholder found in footer (footers are scalar-only)
- `not_entire_paragraph`: block placeholder mixed with other text in the same paragraph
- `ambiguous_placeholder`: placeholder appears multiple times where replacement would be unsafe/ambiguous
- `empty_content`: block field is missing, `None`, or empty string

Skipped placeholders remain unchanged in the output document.

---

## Fallback events

Fallbacks are **intentional safety mechanisms** that prioritize content preservation over structure. Examples:
- `tree_parse_failed`: tree renderer failed; fell back to block renderer
- `markdown_parse_failed`: markdown parsing failed; fell back to plain text
- `nested_table_detected`: nested table detected inside a cell; rendered as a safe text grid
- `table_rendering_failed`: table rendering failed; fell back to text representation

---

## Logs & output files

When `enable_export_trace=True`, `export_to_word()` returns:
- `export_summary`: the in-memory summary as a dict
- `export_summary_log_path`: path to the saved JSON summary file

Additionally, a structured log file is written alongside the JSON summary (same timestamped directory).

### Log directory structure

Export summaries are saved to timestamped directories:

```text
{EXPORT_LOG_DIR}/logs/
  20240124_143022/
    export_summary_output.json
    export_trace_output.log
```

- Each export run gets its own timestamped directory
- Multiple exports in the same second may share the same directory

---

## Configuration

### Enable/Disable observability

Observability is **enabled by default** (`enable_export_trace=True`). To disable:

```python
request = WordExportRequest(
    scalar_fields={...},
    block_fields={...},
    enable_export_trace=False,
)
```

When disabled, no summary is collected and no log files are created (near-zero overhead).

### Environment variables

- `OUTPUT_DIRECTORY`: base directory for output files (relative `output_path` resolved relative to this if set)
- `USE_MARKDOWN_TREE_RENDERER`: enable/disable tree renderer (`true` by default; set `"false"`, `"0"`, or `"no"` to disable)
- `EXPORT_LOG_DIR`: base directory for export logs. If not set, logs are saved under the output directory.

---

## Testing guidance

- Run the test suite: `python -m pytest`
- Inspect generated artifacts under `tests/artifacts/` (used by multiple tests) when debugging output differences.

# Observability Guide: Export Summary and Audit Trail

This guide explains how to access and use the observability features (logs, trace, and export summary) in the deterministic DOCX export library.

## Quick Answer: Where to See It

### 1. **Log Directory** (Automatically Saved)

When `enable_export_trace=True` (default), the export summary is **automatically saved** to a timestamped subdirectory under `logs/`:

```
{EXPORT_LOG_DIR}/logs/
  20240124_143022/
    export_summary_output.json
    export_trace_output.log
```

**Location:** 
- JSON Summary: `{EXPORT_LOG_DIR}/logs/{timestamp}/export_summary_{filename}.json`
- Structured Logs: `{EXPORT_LOG_DIR}/logs/{timestamp}/export_trace_{filename}.log`

- If `EXPORT_LOG_DIR` environment variable is set, use `{EXPORT_LOG_DIR}/logs/{timestamp}/`
- Otherwise, use `{output_dir}/logs/{timestamp}/`
- Each export run gets its own timestamped directory (e.g., `20240124_143022/`)
- The directory structure is created automatically if it doesn't exist

The log file path is also returned in the result: `result["export_summary_log_path"]`

### 2. **Direct Function Call** (Python Script)

When you call `export_to_word()` directly, the summary is in the return value:

```python
from deterministic_docx_export.models.export_models import WordExportRequest
from deterministic_docx_export.services.word_export_service import export_to_word
from pathlib import Path
import json

# Create request (enable_export_trace=True by default)
request = WordExportRequest(
    scalar_fields={"title": "My Doc"},
    block_fields={"summary": "# Summary\n\nContent..."}
)

# Call export
result = export_to_word(
    template_path=Path("template.docx"),
    request=request,
    markdown_mode=True,
    output_path=Path("output.docx")
)

# ? THE SUMMARY IS HERE:
if "export_summary" in result:
    summary = result["export_summary"]
    print(json.dumps(summary, indent=2))
    
    # Also get the log file path
    if "export_summary_log_path" in result:
        print(f"Log saved to: {result['export_summary_log_path']}")
```

**Location:** `result["export_summary"]` and `result["export_summary_log_path"]`

### 3. **API Endpoint** (HTTP Response)

When calling the API endpoint, the summary is in the JSON response:

```bash
# Call the API
curl -X POST "http://localhost:8000/api/brief/{document_id}/summary/export-word-library" \
  -H "Content-Type: application/json" \
  -d '{
    "scalar_fields": {"title": "My Doc"},
    "block_fields": {"summary": "# Summary\n\nContent..."},
    "enable_export_trace": true
  }'
```

**Response includes:**
```json
{
  "success": true,
  "document_id": "...",
  "output_directory": "...",
  "word_file_path": "...",
  "markdown_files": [...],
  "export_summary": {
    "scalar_replacements": [...],
    "block_replacements": [...],
    "skipped_items": [...],
    "fallback_events": [...],
    "warnings": [...]
  },
  "export_summary_log_path": "/path/to/logs/20240124_143022/export_summary_output.json"
}
```

**Location:** Response JSON ? `export_summary` field and `export_summary_log_path` field

### 4. **Structured Log Files** (Automatically Saved)

When `enable_export_trace=True` (default), structured logs are **automatically saved** to a log file in the same timestamped directory as the JSON summary:

```
{EXPORT_LOG_DIR}/logs/
  20240124_143022/
    export_summary_output.json
    export_trace_output.log
```

**Location:** `{EXPORT_LOG_DIR}/logs/{timestamp}/export_trace_{filename}.log`

The log file contains all structured INFO-level messages from the export process, including:
- Scalar placeholder replacements
- Block placeholder replacements
- Skipped items
- Fallback events
- Summary statistics

**Note:** Structured logs are also written to console/terminal if Python logging is configured (see below).

### 5. **Console Logs** (Terminal Output - Optional)

Structured logs also appear in your console/terminal when logging is configured:

```python
import logging
logging.basicConfig(level=logging.INFO)

# When you run export_to_word(), you'll see:
# INFO - Export trace enabled: collecting summary and structured logs
# INFO - Replaced scalar placeholder {{title}} in BODY FLOW (preview: My Doc)
# INFO - Replaced block placeholder {{summary}} in BODY FLOW using block renderer...
# INFO - Export summary: 1 scalar replacements, 1 block replacements, 0 skipped items...
```

**Location:** Console/terminal output (INFO level logs) - only if logging is configured

## Export Summary Structure

The `export_summary` contains:

```json
{
  "scalar_replacements": [
    {
      "token": "{{title}}",
      "region": "BODY",
      "container": "FLOW",
      "value_preview": "My Document Title"
    }
  ],
  "block_replacements": [
    {
      "token": "{{summary}}",
      "region": "BODY",
      "container": "FLOW",
      "block_types": ["heading", "paragraph", "bullet_list"]
    }
  ],
  "skipped_items": [
    {
      "token": "{{summary}}",
      "region": "HEADER",
      "container": "FLOW",
      "reason": "block_in_header"
    }
  ],
  "fallback_events": [
    {
      "event_type": "tree_parse_failed",
      "location": "{{summary}}",
      "reason": "Tree parsing failed, using block-based renderer fallback"
    }
  ],
  "warnings": [
    {
      "message": "Block placeholder {{summary}} found inside TEXTBOX...",
      "context": {"token": "{{summary}}", "region": "BODY", "container": "TEXTBOX"}
    }
  ]
}
```

## Environment Variables

### `EXPORT_LOG_DIR`

Controls where export summary logs are saved:

- **If set**: Logs are saved to `{EXPORT_LOG_DIR}/logs/{timestamp}/`
- **If not set**: Logs are saved to `{output_dir}/logs/{timestamp}/`

**Example:**
```bash
# Set in .env file or environment
export EXPORT_LOG_DIR=/path/to/logs

# Or in .env file:
EXPORT_LOG_DIR=/path/to/logs
```

The library automatically loads `.env` files if `python-dotenv` is available.

## Log Directory Structure

Each export run creates a timestamped directory:

```
logs/
  20240124_140853/
    export_summary_test_basic_blocks.json
    export_trace_test_basic_blocks.log
  20240124_140900/
    export_summary_test_basic_blocks.json
    export_trace_test_basic_blocks.log
    export_summary_test_basic_scalar.json
    export_trace_test_basic_scalar.log
  20240124_140904/
    export_summary_test_complex_nested_markdown.json
    export_trace_test_complex_nested_markdown.log
```

**Benefits:**
- Organized by run time: Easy to find logs from a specific export run
- Grouped together: Multiple exports in the same run are grouped in one directory
- Clean structure: No timestamp duplication in filenames
- Complete audit trail: Both JSON summary and structured log file in same directory
- Easy cleanup: Can delete old timestamp directories to clean up old logs

## Complete Example

```python
import json
import logging
from pathlib import Path
from deterministic_docx_export.models.export_models import WordExportRequest
from deterministic_docx_export.services.word_export_service import export_to_word

# Enable logging to see structured logs
logging.basicConfig(level=logging.INFO, format='%(levelname)s - %(message)s')

# Create request
request = WordExportRequest(
    scalar_fields={"title": "Test Document"},
    block_fields={"summary": "# Summary\n\nTest content"}
)

# Export
result = export_to_word(
    template_path=Path("your_template.docx"),
    request=request,
    markdown_mode=True,
    output_path=Path("output.docx")
)

# ? VIEW THE SUMMARY HERE:
print("\n" + "="*60)
print("EXPORT SUMMARY (AUDIT TRAIL)")
print("="*60)

if "export_summary" in result:
    summary = result["export_summary"]
    
    # Print formatted summary
    print(f"\nScalar Replacements: {len(summary['scalar_replacements'])}")
    for r in summary['scalar_replacements']:
        print(f"  - {r['token']} ? {r['value_preview']}")
    
    print(f"\nBlock Replacements: {len(summary['block_replacements'])}")
    for r in summary['block_replacements']:
        print(f"  - {r['token']} ? {', '.join(r['block_types'])}")
    
    print(f"\nSkipped Items: {len(summary['skipped_items'])}")
    for s in summary['skipped_items']:
        print(f"  - {s['token']}: {s['reason']}")
    
    print(f"\nFallback Events: {len(summary['fallback_events'])}")
    for f in summary['fallback_events']:
        print(f"  - {f['event_type']}: {f.get('reason', 'N/A')}")
    
    print(f"\nWarnings: {len(summary['warnings'])}")
    for w in summary['warnings']:
        print(f"  - {w['message']}")
    
    # Access log file path
    if "export_summary_log_path" in result:
        print(f"\n?? Log file: {result['export_summary_log_path']}")
    
    # Save to file
    with open("export_audit.json", "w") as f:
        json.dump(summary, f, indent=2)
    print("\n? Full summary saved to: export_audit.json")
else:
    print("\n??  No summary (enable_export_trace=False)")
```

## What Gets Logged

### Structured Logs (INFO level, gated by `enable_export_trace`):

1. **Scalar Replacements:**
   ```
   INFO - Replaced scalar placeholder {{document_id}} in BODY FLOW (preview: DOC-12345)
   ```

2. **Block Replacements:**
   ```
   INFO - Replaced block placeholder {{summary}} in BODY FLOW using block renderer (block types: heading, paragraph, bullet_list)
   INFO - Replaced block placeholder {{summary}} in BODY FLOW using tree renderer (block types: heading, paragraph)
   ```

3. **Skipped Items:**
   ```
   INFO - Skipped block placeholder {{summary}} in HEADER/FLOW: block_in_header
   INFO - Skipped block placeholder {{summary}} in BODY/TEXTBOX: block_in_textbox
   ```

4. **Fallback Events:**
   ```
   INFO - Fallback: Tree parsing failed for {{summary}}, falling back to block-based rendering
   INFO - Fallback: Nested table rendering failed for {{table}}, using text-grid
   ```

5. **Summary Log (at end):**
   ```
   INFO - Export summary: 2 scalar replacements, 1 block replacements, 1 skipped items, 1 fallback events, 1 warnings
   INFO - Export summary saved to log directory: /path/to/logs/20240124_143022/export_summary_output.json
   ```

### Warning Logs (always logged, not gated):

- Block placeholder in textbox/header/footer (safety warnings)
- Block placeholder not occupying entire paragraph
- Tree parsing failures (also recorded in summary)

## Enable/Disable Observability

Observability is **enabled by default** but can be disabled:

```python
# Disable observability (no summary, no structured logs)
request = WordExportRequest(
    scalar_fields={"title": "My Doc"},
    block_fields={"summary": "Content..."},
    enable_export_trace=False  # Disable observability
)

result = export_to_word(...)
# result will NOT contain "export_summary" key
# No structured logs will be emitted
# No log files will be created
```

**When Disabled:**
- No `export_summary` in return value
- No structured INFO logs (only warnings remain)
- No log files created
- Near-zero overhead

## Testing

To verify observability is working:

```bash
cd deterministic-docx-export
python -m pytest tests/ -v
```

The tests verify that:
- Summary is collected when `enable_export_trace=True`
- Summary is NOT collected when `enable_export_trace=False`
- All structured logs are properly gated
- No rendering behavior is affected
- Log files are created in the correct location

## Summary

- **Export Summary**: Access via `result["export_summary"]` (when `enable_export_trace=True`)
- **JSON Summary File**: Automatically saved to `{EXPORT_LOG_DIR}/logs/{timestamp}/export_summary_{filename}.json`
- **Structured Log File**: Automatically saved to `{EXPORT_LOG_DIR}/logs/{timestamp}/export_trace_{filename}.log`
- **Log Path**: Available in `result["export_summary_log_path"]` (points to JSON file)
- **Console Logs**: Configure Python logging to see INFO-level messages in console (optional)
- **Enable/Disable**: Use `enable_export_trace` flag in `WordExportRequest` (default: `True`)
- **Environment Variable**: Set `EXPORT_LOG_DIR` to control log location
- **Zero Overhead**: When disabled, observability has near-zero performance impact
- **No Behavior Changes**: Observability never affects rendering or output
