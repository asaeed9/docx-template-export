"""
Compatibility shim for deterministic_docx_export.services.

This module re-exports everything from docx_template_export.services.
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
from docx_template_export.services import *

# Make submodules importable: from deterministic_docx_export.services.word_export_service import ...
import docx_template_export.services.word_export_service as word_export_service
import docx_template_export.services.output_path as output_path

sys.modules['deterministic_docx_export.services.word_export_service'] = word_export_service
sys.modules['deterministic_docx_export.services.output_path'] = output_path
