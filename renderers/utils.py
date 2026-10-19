"""
renderers/utils.py - Re-exports text, formatting, and measurement utilities from core.utils.
"""

from core.utils import (
    natural_sort_key,
    extract_question_prefix_and_body,
    split_line_into_option_items,
    apply_title_case_to_text,
    apply_sentence_case_to_text,
    SUPPORTED_IMAGE_EXTENSIONS,
    ensure_word_compatible_image,
    cm_to_pt,
    pt_to_cm,
    COLOR_NAME_TO_RGB,
    HIGHLIGHT_NAME_TO_INDEX,
    parse_color_to_rgb_int,
    GdiTextMeasurer,
    get_gdi_text_measurer,
)

__all__ = [
    "natural_sort_key",
    "extract_question_prefix_and_body",
    "split_line_into_option_items",
    "apply_title_case_to_text",
    "apply_sentence_case_to_text",
    "SUPPORTED_IMAGE_EXTENSIONS",
    "ensure_word_compatible_image",
    "cm_to_pt",
    "pt_to_cm",
    "COLOR_NAME_TO_RGB",
    "HIGHLIGHT_NAME_TO_INDEX",
    "parse_color_to_rgb_int",
    "GdiTextMeasurer",
    "get_gdi_text_measurer",
]
