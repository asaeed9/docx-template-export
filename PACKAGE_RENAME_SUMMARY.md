# Package Rename Summary: deterministic-docx-export ? docx-template-export

## Overview

**⚠️ Important**: The package `deterministic-docx-export` is frozen at version 0.3.0, which is the final release. The project has been renamed and moved to `docx-template-export`. No version 0.3.1 exists for `deterministic-docx-export`. This is a **pure branding change** - all functionality, APIs, and behavior remain identical.

## What Changed

### Package Name
- **Old**: `deterministic-docx-export`
- **New**: `docx-template-export`

### Installation
- **Old**: `pip install deterministic-docx-export`
- **New**: `pip install docx-template-export`

### Imports
- **Old**: `from deterministic_docx_export import ...`
- **New**: `from docx_template_export import ...`

## Backward Compatibility

? **The old package name continues to work** via a compatibility shim:
- All imports from `deterministic_docx_export` still function
- Deprecation warnings are emitted when using the old package name
- Both packages point to the same code (no duplication)
- Legacy package (`deterministic-docx-export`) is frozen at version 0.3.0 (final release)
- New package (`docx-template-export`) starts at version 0.3.1

## Files Updated

### Documentation
- ? `README.md` - Updated installation instructions and added deprecation note
- ? `CHANGELOG.md` - Updated to indicate v0.3.0 is final release (no v0.3.1 section)
- ? `MIGRATION_GUIDE.md` - Created comprehensive migration guide (NEW)

### Code
- ? `setup.py` - Updated package name to `docx-template-export`
- ? All source files - Updated imports to use `docx_template_export`
- ? All test files - Updated imports to use `docx_template_export`
- ? All example files - Updated imports to use `docx_template_export`

### Compatibility Shim
- ? `deterministic_docx_export/__init__.py` - Main compatibility shim
- ? `deterministic_docx_export/models/__init__.py` - Models shim
- ? `deterministic_docx_export/services/__init__.py` - Services shim
- ? `deterministic_docx_export/renderers/__init__.py` - Renderers shim

## Verification

All tests pass:
- ? Package structure verified
- ? New package imports work
- ? Old package shim works
- ? Object identity verified (both packages point to same objects)
- ? All test suites pass (27+ tests)

## Next Steps for Users

1. **Install the new package**: `pip install docx-template-export`
2. **Update imports**: Change `from deterministic_docx_export` to `from docx_template_export`
3. **Update requirements files**: Update `requirements.txt`, `setup.py`, etc.
4. **See [MIGRATION_GUIDE.md](MIGRATION_GUIDE.md)** for detailed migration instructions

## Next Steps for Maintainers

1. ? Clean old distribution files: `rm -rf dist/ build/ *.egg-info/`
2. ? Build new distribution: `python3 -m build`
3. ? Test locally: `pip install dist/docx_template_export-0.3.1-py3-none-any.whl`
4. ? Upload to PyPI: `python3 -m twine upload dist/*`
5. ? Create GitHub release tag: `git tag v0.3.1 && git push origin v0.3.1`
6. ? Announce the rename (if applicable)

## Important Notes

- **No functional changes**: All APIs, behavior, and output remain identical
- **Legacy package version**: Frozen at 0.3.0 (final release)
- **New package version**: Starts at 0.3.1
- **GitHub URL**: Updated to `https://github.com/asaeed9/deterministic-docx-export`
- **PyPI package**: New package name must be available on PyPI

## Support

For migration help, see:
- [MIGRATION_GUIDE.md](MIGRATION_GUIDE.md) - Comprehensive migration instructions
- [README.md](README.md) - Updated usage examples
- [CHANGELOG.md](CHANGELOG.md) - Detailed technical changelog
