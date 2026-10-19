"""
uln_compiler.py - Backward-compatible facade for ULNCompiler.
Re-exports the core compiler implementation from core.compiler.
"""

from core.compiler import (
    ULNCompiler,
    extract_raw_uln,
    embed_raw_uln_docx,
    scan_folder_for_uln_docx,
)

__all__ = [
    "ULNCompiler",
    "extract_raw_uln",
    "embed_raw_uln_docx",
    "scan_folder_for_uln_docx",
]
