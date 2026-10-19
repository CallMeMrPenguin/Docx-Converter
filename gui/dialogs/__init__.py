"""
gui/dialogs/__init__.py - Re-export dialog modules.
"""

from gui.dialogs.prompt_editor import open_prompt_editor_dialog
from gui.dialogs.update_modal import show_update_modal_dialog
from gui.dialogs.docx_picker import open_docx_picker_dialog
from gui.dialogs.image_preview import (
    load_and_scale_image,
    open_image_preview_dialog,
    open_exercise_pic_match_dialog,
)

__all__ = [
    "open_prompt_editor_dialog",
    "show_update_modal_dialog",
    "open_docx_picker_dialog",
    "load_and_scale_image",
    "open_image_preview_dialog",
    "open_exercise_pic_match_dialog",
]
