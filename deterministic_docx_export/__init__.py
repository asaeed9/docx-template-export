"""
Compatibility shim for deterministic-docx-export package.

This module provides backward compatibility for code that imports from
deterministic_docx_export. The package has been renamed to docx_template_export.

All imports are re-exported from the new package location.
"""

import warnings

# Emit deprecation warning
warnings.warn(
    "The package 'deterministic-docx-export' has been renamed to 'docx-template-export'. "
    "Please update your imports from 'deterministic_docx_export' to 'docx_template_export'. "
    "The old package name will be removed in a future version.",
    DeprecationWarning,
    stacklevel=2
)

# Re-export everything from the new package
from docx_template_export import *
from docx_template_export import __version__

# Re-export submodules for direct imports like:
# from deterministic_docx_export.models import ...
# from deterministic_docx_export.services import ...
# from deterministic_docx_export.renderers import ...
from docx_template_export import models
from docx_template_export import services
from docx_template_export import renderers

__all__ = [
    "__version__",
    "models",
    "services",
    "renderers",
]
