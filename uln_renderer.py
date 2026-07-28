import os
import math
from typing import List, Dict, Any, Optional
from uln_parser import ULNBlock, InlineSpan, PicInfo

try:
    import pythoncom
    import win32com.client
    pywin32_available = True
except ImportError:
    pywin32_available = False

def cm_to_pt(cm: float) -> float:
    return float(cm) * 28.346456692913385

# Color converter helper for MS Word COM (RGB integer: R + (G * 256) + (B * 65536))
COLOR_NAME_TO_RGB = {
    "red": 255,                 # RGB(255, 0, 0)
    "blue": 16711680,           # RGB(0, 0, 255)
    "green": 32768,             # RGB(0, 128, 0)
    "yellow": 65535,            # RGB(255, 255, 0)
    "purple": 8388736,          # RGB(128, 0, 128)
    "orange": 42495,            # RGB(255, 165, 0)
    "black": 0,
    "white": 16777215,
    "darkblue": 9125196,        # RGB(12, 15, 139)
    "grey": 8421504,
    "gray": 8421504,
}

HIGHLIGHT_NAME_TO_INDEX = {
    "yellow": 7,    # wdYellow = 7
    "green": 4,     # wdGreen = 4
    "cyan": 3,      # wdTurquoise = 3
    "pink": 5,      # wdPink = 5
    "blue": 9,      # wdBlue = 9
    "red": 6,       # wdRed = 6
    "gray": 15,     # wdGray25 = 15
}

def parse_color_to_rgb_int(color_str: str) -> Optional[int]:
    if not color_str:
        return None
    clean = color_str.strip().lower()
    if clean in COLOR_NAME_TO_RGB:
        return COLOR_NAME_TO_RGB[clean]
    if clean.startswith('#'):
        hex_val = clean.lstrip('#')
        if len(hex_val) == 6:
            r = int(hex_val[0:2], 16)
            g = int(hex_val[2:4], 16)
            b = int(hex_val[4:6], 16)
            return r + (g * 256) + (b * 65536)
    return None


class ULNWordRenderer:
    """Renderer that executes pywin32 COM automation commands to construct Word (.docx) documents."""

    def __init__(self, settings: Optional[Dict[str, Any]] = None):
        self.settings = settings or {}

        # Default Settings
        self.font_name = self.settings.get("font_name", "Times New Roman")
        self.font_size = float(self.settings.get("font_size", 12.0))
        self.margin_top = float(self.settings.get("margin_top", 2.0))
        self.margin_bottom = float(self.settings.get("margin_bottom", 2.0))
        self.margin_left = float(self.settings.get("margin_left", 3.0))
        self.margin_right = float(self.settings.get("margin_right", 1.5))
        self.line_spacing = float(self.settings.get("line_spacing", 1.15))
        self.enable_page_numbers = self.settings.get("enable_page_numbers", True)

    def configure_document(self, doc):
        """Applies page setup margins and optional page numbering."""
        ps = doc.PageSetup
        ps.PageWidth = cm_to_pt(21.0)
        ps.PageHeight = cm_to_pt(29.7)
        ps.TopMargin = cm_to_pt(self.margin_top)
        ps.BottomMargin = cm_to_pt(self.margin_bottom)
        ps.LeftMargin = cm_to_pt(self.margin_left)
        ps.RightMargin = cm_to_pt(self.margin_right)

        if self.enable_page_numbers:
            for section in doc.Sections:
                footer = section.Footers(1)  # wdHeaderFooterPrimary = 1
                footer.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter = 1
                footer.Range.Font.Name = self.font_name
                footer.Range.Font.Size = 10
                footer.Range.Text = ""

                sel_range = footer.Range
                doc.Fields.Add(Range=sel_range, Type=-1, Text="PAGE")  # wdFieldPage = -1
                footer.Range.InsertAfter(" / ")
                end_range = footer.Range
                end_range.Collapse(0)  # wdCollapseEnd = 0
                doc.Fields.Add(Range=end_range, Type=-1, Text="NUMPAGES")

    def write_inline_spans(self, sel, spans: List[InlineSpan], default_bold: bool = False, default_italic: bool = False):
        """Writes formatted text runs to MS Word selection."""
        for span in spans:
            sel.Font.Name = self.font_name
            sel.Font.Size = self.font_size
            sel.Font.Bold = 1 if (span.bold or default_bold) else 0
            sel.Font.Italic = 1 if (span.italic or default_italic) else 0
            sel.Font.Underline = 1 if span.underline else 0

            # Text color
            if span.color:
                rgb_int = parse_color_to_rgb_int(span.color)
                if rgb_int is not None:
                    try:
                        sel.Font.Color = rgb_int
                    except Exception:
                        pass
            else:
                try:
                    sel.Font.ColorIndex = 0  # wdAuto
                except Exception:
                    pass

            # Background Highlight
            if span.bg_color:
                hl_idx = HIGHLIGHT_NAME_TO_INDEX.get(span.bg_color.lower(), 7)
                try:
                    sel.Font.HighlightColorIndex = hl_idx
                except Exception:
                    pass
            else:
                try:
                    sel.Font.HighlightColorIndex = 0  # wdNoHighlight = 0
                except Exception:
                    pass

            text = span.text.upper() if span.uppercase else span.text
            sel.TypeText(text)

            # Reset Highlight to prevent leak
            try:
                sel.Font.HighlightColorIndex = 0
            except Exception:
                pass

    def render(self, blocks: List[ULNBlock], doc, word):
        """Renders parsed ULNBlocks into the active document."""
        self.configure_document(doc)
        sel = word.Selection

        printable_width_cm = 21.0 - self.margin_left - self.margin_right
        printable_width_pt = cm_to_pt(printable_width_cm)

        for block in blocks:
            tag = block.tag

            if tag == "H1":
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.SpaceBefore = 18
                sel.ParagraphFormat.SpaceAfter = 12
                sel.ParagraphFormat.KeepWithNext = True
                sel.ParagraphFormat.Alignment = 0  # Left
                self.write_inline_spans(sel, block.spans, default_bold=True)
                sel.TypeParagraph()

            elif tag == "H2":
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.SpaceBefore = 14
                sel.ParagraphFormat.SpaceAfter = 8
                sel.ParagraphFormat.KeepWithNext = True
                sel.ParagraphFormat.Alignment = 0
                self.write_inline_spans(sel, block.spans, default_bold=True)
                sel.TypeParagraph()

            elif tag == "H3":
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.SpaceBefore = 10
                sel.ParagraphFormat.SpaceAfter = 6
                sel.ParagraphFormat.KeepWithNext = True
                sel.ParagraphFormat.Alignment = 0
                self.write_inline_spans(sel, block.spans, default_bold=True)
                sel.TypeParagraph()

            elif tag in ["P0", "P"]:
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.SpaceBefore = 6
                sel.ParagraphFormat.SpaceAfter = 4
                sel.ParagraphFormat.KeepWithNext = False
                sel.ParagraphFormat.Alignment = 0
                self.write_inline_spans(sel, block.spans)
                sel.TypeParagraph()

            elif tag == "P1":
                sel.ParagraphFormat.LeftIndent = cm_to_pt(0.5)
                sel.ParagraphFormat.SpaceBefore = 4
                sel.ParagraphFormat.SpaceAfter = 3
                sel.ParagraphFormat.KeepWithNext = False
                sel.ParagraphFormat.Alignment = 0
                self.write_inline_spans(sel, block.spans)
                sel.TypeParagraph()

            elif tag == "P2":
                sel.ParagraphFormat.LeftIndent = cm_to_pt(1.0)
                sel.ParagraphFormat.SpaceBefore = 3
                sel.ParagraphFormat.SpaceAfter = 3
                sel.ParagraphFormat.KeepWithNext = False
                sel.ParagraphFormat.Alignment = 0
                self.write_inline_spans(sel, block.spans)
                sel.TypeParagraph()

            elif tag == "TAB2":
                # 2-Column side-by-side using Tab Stops
                col_width_cm = printable_width_cm / 2.0
                tab_pos_cm = col_width_cm

                sel.ParagraphFormat.LeftIndent = cm_to_pt(0.5)
                sel.ParagraphFormat.SpaceBefore = 3
                sel.ParagraphFormat.SpaceAfter = 3
                sel.ParagraphFormat.TabStops.ClearAll()
                sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(tab_pos_cm), Alignment=0)

                # Write Column 1
                self.write_inline_spans(sel, block.col1_spans)

                # Tab over to Column 2
                sel.TypeText("\t")

                # Write Column 2 or handle Pic if embedded
                if block.pic:
                    self.render_pic(sel, doc, block.pic)
                else:
                    self.write_inline_spans(sel, block.col2_spans)

                sel.TypeParagraph()
                sel.ParagraphFormat.TabStops.ClearAll()

            elif tag == "BOX":
                self.render_box(sel, doc, block, printable_width_cm)

            elif tag == "QUOTE":
                sel.ParagraphFormat.LeftIndent = cm_to_pt(1.0)
                sel.ParagraphFormat.RightIndent = cm_to_pt(1.0)
                sel.ParagraphFormat.SpaceBefore = 6
                sel.ParagraphFormat.SpaceAfter = 6
                sel.ParagraphFormat.Alignment = 0
                self.write_inline_spans(sel, block.spans, default_italic=True)
                sel.TypeParagraph()
                sel.ParagraphFormat.RightIndent = 0

            elif tag == "PIC":
                if block.pic:
                    sel.ParagraphFormat.LeftIndent = 0
                    sel.ParagraphFormat.SpaceBefore = 6
                    sel.ParagraphFormat.SpaceAfter = 6
                    if block.pic.pos == "center":
                        sel.ParagraphFormat.Alignment = 1  # Center
                    else:
                        sel.ParagraphFormat.Alignment = 0  # Left

                    self.render_pic(sel, doc, block.pic)
                    sel.TypeParagraph()

    def render_pic(self, sel, doc, pic: PicInfo):
        """Renders an image file if available, or a clean visually framed diagram placeholder."""
        if pic.filepath and os.path.exists(pic.filepath):
            try:
                shape = sel.InlineShapes.AddPicture(FileName=os.path.abspath(pic.filepath))
                # Adjust scale
                if pic.size == "small":
                    shape.Width = cm_to_pt(4.0)
                elif pic.size == "large":
                    shape.Width = cm_to_pt(12.0)
                else:
                    shape.Width = cm_to_pt(7.0)
                return
            except Exception as e:
                print(f"[ULNRenderer] Warning adding picture {pic.filepath}: {e}")

        # Visual Placeholder Box for OCR / Diagram References
        desc_text = f"🖼️ [ DIAGRAM / IMAGE: {pic.description} ]"
        sel.Font.Name = self.font_name
        sel.Font.Size = 10.0
        sel.Font.Italic = True
        sel.Font.Bold = True
        try:
            sel.Font.Color = 8421504  # Grey
        except Exception:
            pass
        sel.TypeText(desc_text)

    def render_box(self, sel, doc, block: ULNBlock, printable_width_cm: float):
        """Renders framed Word Bank / Rule callout box with border styling."""
        box_table = doc.Tables.Add(Range=sel.Range, NumRows=1, NumColumns=1)
        try:
            box_table.Rows.Alignment = 1  # Center
        except Exception:
            pass
        box_table.Columns(1).Width = cm_to_pt(printable_width_cm)

        cell = box_table.Cell(1, 1)
        cell.VerticalAlignment = 1

        # Border styling: Single line 1pt black border
        for border_id in [-1, -2, -3, -4]:  # Top, Left, Bottom, Right
            try:
                cell.Borders(border_id).LineStyle = 1
                cell.Borders(border_id).LineWidth = 8  # 1pt
                cell.Borders(border_id).Color = 0  # Black
            except Exception:
                pass

        p = cell.Range
        p.ParagraphFormat.SpaceBefore = 4
        p.ParagraphFormat.SpaceAfter = 4
        p.ParagraphFormat.LeftIndent = cm_to_pt(0.3)
        p.ParagraphFormat.Alignment = 0

        # Select inside cell and write spans
        sel.Start = cell.Range.Start
        self.write_inline_spans(sel, block.spans)

        # Move selection past table
        sel.Start = box_table.Range.End
        sel.TypeParagraph()
