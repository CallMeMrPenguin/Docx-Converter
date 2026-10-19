"""
uln_parser.py - Backward-compatible facade for ULNParser and AST models.
Re-exports from core.models and core.parser.
"""

from core.models import (
    InlineSpan,
    PicInfo,
    ULNTableCell,
    ULNTableRow,
    ULNTableData,
    ULNBlock,
)
from core.parser import (
    ULNParser,
    parse_pic_tag,
    parse_inline_spans,
    split_bracket_safe_columns,
)

__all__ = [
    "InlineSpan",
    "PicInfo",
    "ULNTableCell",
    "ULNTableRow",
    "ULNTableData",
    "ULNBlock",
    "ULNParser",
    "parse_pic_tag",
    "parse_inline_spans",
    "split_bracket_safe_columns",
]
