"""
Compatibility shim for deterministic_docx_export.renderers.

This module re-exports everything from docx_template_export.renderers.
"""

import warnings
import sys

warnings.warn(
    "The package 'deterministic-docx-export' has been renamed to 'docx-template-export'. "
    "Please update your imports from 'deterministic_docx_export' to 'docx_template_export'. "
    "The old package name will be removed in a future version.",
    DeprecationWarning,
    stacklevel=2
)

# Import and re-export everything from the new package
from docx_template_export.renderers import *

# Make submodules importable: from deterministic_docx_export.renderers.tree_renderer import ...
import docx_template_export.renderers.tree_renderer as tree_renderer

sys.modules['deterministic_docx_export.renderers.tree_renderer'] = tree_renderer
