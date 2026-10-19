"""
gui_image_preview.py - Backward-compatible facade for image preview and matching dialogs.
Re-exports from gui.dialogs.image_preview.
"""

from gui.dialogs.image_preview import (
    load_and_scale_image,
    open_image_preview_dialog,
    open_exercise_pic_match_dialog,
    extract_pic_questions_from_uln,
)

__all__ = [
    "load_and_scale_image",
    "open_image_preview_dialog",
    "open_exercise_pic_match_dialog",
    "extract_pic_questions_from_uln",
]
