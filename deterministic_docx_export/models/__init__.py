"""
Compatibility shim for deterministic_docx_export.models.

This module re-exports everything from docx_template_export.models.
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
from docx_template_export.models import *

# Make submodules importable: from deterministic_docx_export.models.export_models import ...
# This is done by importing the modules and adding them to sys.modules
import docx_template_export.models.export_models as export_models
import docx_template_export.models.export_summary as export_summary
import docx_template_export.models.export_config as export_config
import docx_template_export.models.markdown_tree as markdown_tree

sys.modules['deterministic_docx_export.models.export_models'] = export_models
sys.modules['deterministic_docx_export.models.export_summary'] = export_summary
sys.modules['deterministic_docx_export.models.export_config'] = export_config
sys.modules['deterministic_docx_export.models.markdown_tree'] = markdown_tree
