from renderers.common import (
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
)

from renderers.exercises import (
    McqOptionsRendererMixin,
    TableRendererMixin,
    TabColumnsRendererMixin,
    PicturesRendererMixin,
    BoxesRendererMixin,
    ContainersRendererMixin,
)


class RendererBlocksMixin(
    TypographyMixin,
    NumberingMixin,
    InlineWriterMixin,
    McqOptionsRendererMixin,
    TableRendererMixin,
    TabColumnsRendererMixin,
    PicturesRendererMixin,
    BoxesRendererMixin,
    ContainersRendererMixin,
):
    """
    Composite mixin combining all exercise-specific renderers and shared typography/formatting utilities:
    - Typography & GDI measurement (common/typography.py)
    - Numbering & List formats (common/numbering.py)
    - Inline text spans & styling (common/inline_writer.py)
    - Multiple Choice Options [OPT] (exercises/mcq_options.py)
    - Tables & Side-diagrams [TABLE] (exercises/tables.py)
    - Multi-column Tab Stops [TAB2], [TAB3], [TAB4] (exercises/tab_columns.py)
    - Pictures & Picture Grids [PIC], [PIC_GRID] (exercises/pictures.py)
    - Formulas, Callouts & Word Banks [BOX] (exercises/boxes.py)
    - Auto-numbered containers [NUM] (exercises/containers.py)
    """
    pass


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
