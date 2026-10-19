"""
uln_renderer.py - Root facade for ULNWordRenderer.
Maintains 100% backward compatibility for all existing scripts, builds, and PyInstaller configs.
The implementation lives in renderers/base.py and renderers/.
"""

from renderers.base import ULNWordRenderer
from renderers import (
    RendererBlocksMixin,
    cm_to_pt,
    pt_to_cm,
    parse_color_to_rgb_int,
    parse_highlight_to_index,
    COLOR_NAME_TO_RGB,
    HIGHLIGHT_NAME_TO_INDEX,
    GdiTextMeasurer,
    get_gdi_text_measurer,
    TypographyMixin,
    NumberingMixin,
    InlineWriterMixin,
    McqOptionsRendererMixin,
    TableRendererMixin,
    TabColumnsRendererMixin,
    PicturesRendererMixin,
    BoxesRendererMixin,
    ContainersRendererMixin,
)

__all__ = [
    "ULNWordRenderer",
    "RendererBlocksMixin",
    "cm_to_pt",
    "pt_to_cm",
    "parse_color_to_rgb_int",
    "parse_highlight_to_index",
    "COLOR_NAME_TO_RGB",
    "HIGHLIGHT_NAME_TO_INDEX",
    "GdiTextMeasurer",
    "get_gdi_text_measurer",
    "TypographyMixin",
    "NumberingMixin",
    "InlineWriterMixin",
    "McqOptionsRendererMixin",
    "TableRendererMixin",
    "TabColumnsRendererMixin",
    "PicturesRendererMixin",
    "BoxesRendererMixin",
    "ContainersRendererMixin",
]
