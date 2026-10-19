from dataclasses import dataclass, field
from typing import List, Optional

@dataclass
class InlineSpan:
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    uppercase: bool = False
    color: Optional[str] = None      # e.g., "red", "#FF0000"
    bg_color: Optional[str] = None   # e.g., "yellow", "#FFFF00"
    is_instruction: bool = False     # Flag for [ins]...[/ins] instruction formatting

@dataclass
class PicInfo:
    description: str
    pos: str = "center"   # "inline", "center", "right"
    size: str = "medium"  # "small", "medium", "large"
    filepath: Optional[str] = None

@dataclass
class ULNTableCell:
    content: str
    is_header: bool = False
    spans: List[InlineSpan] = field(default_factory=list)

@dataclass
class ULNTableRow:
    cells: List[ULNTableCell] = field(default_factory=list)
    is_header: bool = False

@dataclass
class ULNTableData:
    rows: List[ULNTableRow] = field(default_factory=list)
    borderless: bool = False

@dataclass
class ULNBlock:
    tag: str  # "H1", "H2", "H3", "H4", "H5", "H6", "P0", "P1", "P2", "TAB", "TAB2", "TAB3", "TAB4", "BOX", "QUOTE", "PIC", "TABLE", "P", "INS"
    content: str = ""
    col1: str = ""
    col2: str = ""
    cols: List[str] = field(default_factory=list)
    col1_spans: List[InlineSpan] = field(default_factory=list)
    col2_spans: List[InlineSpan] = field(default_factory=list)
    cols_spans: List[List[InlineSpan]] = field(default_factory=list)
    spans: List[InlineSpan] = field(default_factory=list)
    pic: Optional[PicInfo] = None
    table_data: Optional[ULNTableData] = None
    children: List['ULNBlock'] = field(default_factory=list)
    is_instruction: bool = False
    sub_tag: Optional[str] = None
