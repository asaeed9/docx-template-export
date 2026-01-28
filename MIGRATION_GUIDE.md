# Migration Guide: deterministic-docx-export ? docx-template-export

This guide helps you migrate from the old package name `deterministic-docx-export` to the new package name `docx-template-export`.

## Overview

**⚠️ Important**: The package `deterministic-docx-export` is frozen at version 0.3.0, which is the final release. The project has been renamed and moved to `docx-template-export`. No version 0.3.1 exists for `deterministic-docx-export`. This is a **pure branding change** - all functionality, APIs, and behavior remain identical.

## Quick Migration

### 1. Update Installation

**Legacy Package (frozen at 0.3.0):**
```bash
pip install deterministic-docx-export==0.3.0
```

**New Package:**
```bash
pip install docx-template-export
```

### 2. Update Imports

**Old:**
```python
from deterministic_docx_export.models.export_models import WordExportRequest
from deterministic_docx_export.services.word_export_service import export_to_word
```

**New:**
```python
from docx_template_export.models.export_models import WordExportRequest
from docx_template_export.services.word_export_service import export_to_word
```

### 3. Update Requirements Files

**requirements.txt:**
```diff
- deterministic-docx-export==0.3.0
+ docx-template-export>=0.3.1
```

**setup.py or pyproject.toml:**
```diff
  install_requires=[
-     "deterministic-docx-export==0.3.0",
+     "docx-template-export>=0.3.1",
  ]
```

## Backward Compatibility

The old package name (`deterministic-docx-export`) will continue to work via a compatibility shim. However:

- ?? **Deprecation warnings** will be emitted when using the old package name
- ? **All functionality remains identical** - both packages point to the same code
- ? **No breaking changes** - existing code will continue to work
- ? **Future removal**: The old package name may be removed in a future major version

### Example with Deprecation Warning

```python
# This still works but shows a deprecation warning
from deterministic_docx_export.models.export_models import WordExportRequest
# DeprecationWarning: The package 'deterministic-docx-export' has been renamed 
# to 'docx-template-export'. Please update your imports...
```

## Step-by-Step Migration

### Step 1: Install the New Package

```bash
pip uninstall deterministic-docx-export
pip install docx-template-export
```

### Step 2: Update Your Code

Use find-and-replace in your codebase:

**Find:**
```
from deterministic_docx_export
import deterministic_docx_export
```

**Replace:**
```
from docx_template_export
import docx_template_export
```

### Step 3: Update Tests

If you have tests that import the package, update them similarly:

```python
# Old
from deterministic_docx_export.models.export_models import WordExportRequest

# New
from docx_template_export.models.export_models import WordExportRequest
```

### Step 4: Verify

Run your tests to ensure everything works:

```bash
pytest
```

## Common Import Patterns

### Direct Imports

```python
# Old
from deterministic_docx_export import __version__
from deterministic_docx_export.models import export_models
from deterministic_docx_export.services import word_export_service

# New
from docx_template_export import __version__
from docx_template_export.models import export_models
from docx_template_export.services import word_export_service
```

### Submodule Imports

```python
# Old
from deterministic_docx_export.models.export_models import WordExportRequest
from deterministic_docx_export.models.export_config import ListRenderConfig
from deterministic_docx_export.models.export_summary import ExportSummary
from deterministic_docx_export.services.word_export_service import export_to_word
from deterministic_docx_export.renderers.tree_renderer import render_markdown_tree_to_docx

# New
from docx_template_export.models.export_models import WordExportRequest
from docx_template_export.models.export_config import ListRenderConfig
from docx_template_export.models.export_summary import ExportSummary
from docx_template_export.services.word_export_service import export_to_word
from docx_template_export.renderers.tree_renderer import render_markdown_tree_to_docx
```

## Version Compatibility

**Note:** `deterministic-docx-export` is frozen at version 0.3.0. The new package `docx-template-export` starts at version 0.3.1.

## FAQ

### Q: Do I need to migrate immediately?

**A:** No, but it's recommended. The old package name will continue to work but will show deprecation warnings. Migrate when convenient.

### Q: Will my existing code break?

**A:** No. The old package name works via a compatibility shim. However, you should migrate to avoid deprecation warnings and future compatibility issues.

### Q: Can I use both package names in the same project?

**A:** Technically yes (they point to the same code), but it's not recommended. Choose one and stick with it.

### Q: What if I'm using a pinned version?

**A:** Update your version pin to use the new package name:

```diff
- deterministic-docx-export==0.3.0
+ docx-template-export>=0.3.1
```

**Note:** `deterministic-docx-export` is frozen at 0.3.0. No version 0.3.1 exists for the legacy package.

### Q: Are there any API changes?

**A:** No. This is purely a package rename. All APIs, function signatures, and behavior remain identical.

## Need Help?

If you encounter any issues during migration:

1. Check that you've installed the new package: `pip list | grep docx-template-export`
2. Verify your imports are updated correctly
3. Run your tests to ensure everything works
4. Check the [README.md](README.md) for usage examples

## Summary

- ? **Install**: `pip install docx-template-export`
- ? **Import**: `from docx_template_export import ...`
- ? **No breaking changes**: All functionality remains identical
- ?? **Old package deprecated**: Migrate when convenient
