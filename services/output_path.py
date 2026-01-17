"""
Output directory path management utilities.

Provides simple path resolution for output directories.
Configurable via OUTPUT_DIRECTORY environment variable (default: "output").
"""
import os
from pathlib import Path

# Configuration from environment variable
OUTPUT_DIRECTORY = os.getenv("OUTPUT_DIRECTORY", "output")


def get_output_base_dir() -> Path:
    """
    Get the base output directory path.
    
    Uses OUTPUT_DIRECTORY environment variable if set, otherwise defaults to "output".
    Resolves to absolute path relative to current working directory.
    
    Returns:
        Path to base output directory (absolute path)
    """
    output_dir = Path(OUTPUT_DIRECTORY)
    if not output_dir.is_absolute():
        output_dir = Path.cwd() / output_dir
    return output_dir.resolve()


def get_output_dir_for_document(document_id: str) -> Path:
    """
    Get output directory path for a document.
    
    Document ID format: "pdf_name/experiment_id"
    Structure: {output_base}/{document_id}/
    
    Args:
        document_id: Document identifier in format "pdf_name/experiment_id"
        
    Returns:
        Path to document output directory (absolute path)
    """
    base_dir = get_output_base_dir()
    output_dir = base_dir / document_id
    return output_dir


def ensure_output_dir_exists(path: Path) -> Path:
    """
    Ensure the output directory exists, creating it if necessary.
    
    Args:
        path: Path to the output directory
        
    Returns:
        The same path (for chaining)
    """
    path.mkdir(parents=True, exist_ok=True)
    return path
