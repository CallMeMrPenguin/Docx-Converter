"""
renderer_blocks.py - Modularized Renderer Mixin Exporter.
Preserves 100% backward compatibility for all existing scripts and GUI modules.
All modular renderers are organized under renderers/exercises/ and renderers/common/.
"""

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
