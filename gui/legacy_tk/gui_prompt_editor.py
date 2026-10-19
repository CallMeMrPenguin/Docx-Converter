"""
gui_prompt_editor.py - Backward-compatible facade for AI prompt editor dialog.
Re-exports from gui.dialogs.prompt_editor.
"""

from gui.dialogs.prompt_editor import (
    get_prompt_storage_path,
    get_default_prompt_path,
    load_prompt_text,
    save_prompt_text,
    reset_prompt_text,
    open_prompt_editor_dialog,
)

__all__ = [
    "get_prompt_storage_path",
    "get_default_prompt_path",
    "load_prompt_text",
    "save_prompt_text",
    "reset_prompt_text",
    "open_prompt_editor_dialog",
]
