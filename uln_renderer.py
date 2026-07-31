import os
import re
import math
from typing import List, Dict, Any, Optional
from uln_parser import ULNBlock, InlineSpan, PicInfo, parse_pic_tag

try:
    import pythoncom
    import win32com.client
    pywin32_available = True
except ImportError:
    pywin32_available = False

def cm_to_pt(cm: float) -> float:
    return float(cm) * 28.346456692913385

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


def split_line_into_option_items(line_text: str) -> List[str]:
    """
    Splits line text into multi-column items automatically by detecting:
    1. Explicit \t characters
    2. Multiple option choices like A. text B. text C. text D. text or a. b. c. d. e. f.
    3. Multiple spaces / double spaces
    """
    if not line_text:
        return []

    if '\t' in line_text:
        return [x.strip() for x in line_text.split('\t') if x.strip()]

    # Search for option choices like A. B. C. D. E. F. or a. b. c. d. e. f. or **a.** **b.**
    pattern = r'(?:^|\s+)(?:(?:\*\*|\*|\[|\(?)*([a-zA-Z][\.\)])(?:\*\*|\*|\]|\}|\{u\}|\))*)(?=\s+|$)'
    matches = list(re.finditer(pattern, line_text))
    if len(matches) >= 2:
        items = []
        first_start = matches[0].start()
        q_prefix = line_text[:first_start].strip()

        for i in range(len(matches)):
            start = matches[i].start()
            end = matches[i + 1].start() if i + 1 < len(matches) else len(line_text)
            item_txt = line_text[start:end].strip()
            if i == 0 and q_prefix:
                item_txt = f"{q_prefix} {item_txt}"
            items.append(item_txt)
        return items

    # Fallback to double space split if multiple columns exist
    if '  ' in line_text:
        parts = [p.strip() for p in re.split(r'\s{2,}', line_text) if p.strip()]
        if len(parts) >= 2:
            return parts

    return [line_text.strip()]


class ULNWordRenderer:
    """Pure tag-driven pywin32 COM document renderer with adaptive column calculation and zero text wrapping into column 1."""

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
        self.space_before = float(self.settings.get("space_before", 4.0))
        self.enable_page_numbers = self.settings.get("enable_page_numbers", True)
        self.user_images = list(self.settings.get("user_images", []))
        self.user_img_idx = 0

    def get_next_image_path(self, pic: Optional[PicInfo] = None) -> Optional[str]:
        """Returns next user-queued image in order, or falls back to test pic directory."""
        if self.user_img_idx < len(self.user_images):
            imgPath = self.user_images[self.user_img_idx]
            self.user_img_idx += 1
            if os.path.exists(imgPath):
                return imgPath

        if pic and pic.filepath and os.path.exists(pic.filepath):
            return pic.filepath

        test_pic_dir = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "test pic"))
        if os.path.exists(test_pic_dir):
            pics = [os.path.join(test_pic_dir, f) for f in os.listdir(test_pic_dir) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
            if pics:
                idx = abs(hash(pic.description if pic else "img")) % len(pics)
                return pics[idx]

        return None

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

    def write_inline_spans(self, sel, spans: List[InlineSpan], default_bold: bool = False, default_italic: bool = False, default_uppercase: bool = False):
        """Writes formatted text runs strictly according to span AST properties."""
        for idx, span in enumerate(spans):
            text = span.text

            # Check if span is an inline [PIC...] tag
            if text.startswith("[PIC:") or text.strip().upper() == "[PIC]":
                pic_info = parse_pic_tag(text) or PicInfo(description="Activity Picture", pos="center", size="medium")
                self.render_pic(sel, None, pic_info)
                try: sel.Font.Underline = 0
                except Exception: pass
                try: sel.Font.HighlightColorIndex = 0
                except Exception: pass
                continue

            sel.Font.Name = self.font_name
            sel.Font.Size = self.font_size
            
            is_bold = span.bold or default_bold
            is_italic = span.italic or default_italic
            is_upper = span.uppercase or default_uppercase

            # Auto-capitalize & bold option letters: "a. ", "b. ", "c. ", "d. ", "e. ", "f. ", "a) ", "b) "
            opt_match = re.match(r'^(\s*(?:\d+\.\s*)?)([a-zA-Z])([\.\)])(\s*.*)$', text)
            if opt_match:
                prefix_num = opt_match.group(1)
                opt_let = opt_match.group(2).upper()
                opt_punct = opt_match.group(3)
                rest_txt = opt_match.group(4)

                if prefix_num:
                    sel.Font.Bold = 1 if default_bold else 0
                    sel.Font.Italic = 1 if default_italic else 0
                    sel.Font.Underline = 0
                    sel.TypeText(prefix_num)

                sel.Font.Bold = 1
                sel.Font.Italic = 0
                sel.Font.Underline = 0
                sel.TypeText(f"{opt_let}{opt_punct} ")

                text = rest_txt
                is_bold = span.bold

            sel.Font.Bold = 1 if is_bold else 0
            sel.Font.Italic = 1 if is_italic else 0
            sel.Font.Underline = 1 if span.underline else 0

            if span.color:
                rgb_int = parse_color_to_rgb_int(span.color)
                if rgb_int is not None:
                    try:
                        sel.Font.Color = rgb_int
                    except Exception:
                        pass
            else:
                try:
                    sel.Font.ColorIndex = 0
                except Exception:
                    pass

            if span.bg_color:
                hl_idx = HIGHLIGHT_NAME_TO_INDEX.get(span.bg_color.lower(), 7)
                try:
                    sel.Font.HighlightColorIndex = hl_idx
                except Exception:
                    pass
            else:
                try:
                    sel.Font.HighlightColorIndex = 0
                except Exception:
                    pass

            # Standardize excessive underscores (>45) to an optimal blank line length
            if re.match(r'^_{30,}$', text):
                text = '_' * 35

            # Replace <blank>, <BLANK>, [BLANK] tag strings with normalized student answer line _____
            text = re.sub(r'<(?:blank|BLANK)>|\[(?:blank|BLANK)\]', '___________', text)

            text = text.upper() if is_upper else text
            sel.TypeText(text)

            try:
                sel.Font.Underline = 0
            except Exception:
                pass
            try:
                sel.Font.HighlightColorIndex = 0
            except Exception:
                pass

    def setup_tab_stops(self, sel, num_cols: int, left_indent_cm: float, printable_width_cm: float) -> float:
        """Calculates exact distance formulas for tab stops so text NEVER wraps unexpectedly."""
        remaining_width_cm = printable_width_cm - left_indent_cm
        col_width_cm = remaining_width_cm / max(1, num_cols)

        sel.ParagraphFormat.LeftIndent = cm_to_pt(left_indent_cm)
        sel.ParagraphFormat.TabStops.ClearAll()

        for c in range(1, num_cols):
            tab_pos_cm = left_indent_cm + (col_width_cm * c)
            sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(tab_pos_cm), Alignment=0)

        return col_width_cm

    def render(self, blocks: List[ULNBlock], doc, word):
        """Renders parsed ULNBlocks into the active document purely driven by structural AST tags."""
        self.configure_document(doc)
        sel = word.Selection

        printable_width_cm = 21.0 - self.margin_left - self.margin_right

        idx_block = 0
        while idx_block < len(blocks):
            block = blocks[idx_block]
            tag = block.tag

            if tag == "H1":
                try:
                    sel.Style = doc.Styles("Heading 1")
                except Exception:
                    pass
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.FirstLineIndent = 0
                sel.ParagraphFormat.SpaceBefore = 14
                sel.ParagraphFormat.SpaceAfter = 6
                sel.ParagraphFormat.KeepWithNext = True
                sel.ParagraphFormat.Alignment = 0  # Left
                self.write_inline_spans(sel, block.spans, default_bold=True, default_uppercase=True)
                sel.TypeParagraph()

            elif tag == "H2":
                try:
                    sel.Style = doc.Styles("Heading 2")
                except Exception:
                    pass
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.FirstLineIndent = 0
                sel.ParagraphFormat.SpaceBefore = 12
                sel.ParagraphFormat.SpaceAfter = 4
                sel.ParagraphFormat.KeepWithNext = True
                sel.ParagraphFormat.Alignment = 0
                self.write_inline_spans(sel, block.spans, default_bold=True, default_uppercase=True)
                sel.TypeParagraph()

            elif tag == "H3":
                try:
                    sel.Style = doc.Styles("Heading 3")
                except Exception:
                    pass
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.FirstLineIndent = 0
                sel.ParagraphFormat.SpaceBefore = 10
                sel.ParagraphFormat.SpaceAfter = 4
                sel.ParagraphFormat.KeepWithNext = True
                sel.ParagraphFormat.Alignment = 0
                self.write_inline_spans(sel, block.spans, default_bold=True, default_uppercase=True)
                sel.TypeParagraph()

            elif tag in ["P0", "P"]:
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.FirstLineIndent = 0
                sel.ParagraphFormat.SpaceBefore = 6
                sel.ParagraphFormat.SpaceAfter = 4
                sel.ParagraphFormat.KeepWithNext = False
                sel.ParagraphFormat.Alignment = 0

                # Check if paragraph is ONLY a standalone blank line _____ or <blank> or [BLANK]
                ends_with_blank = bool(re.search(r'^\s*(?:_{3,}|<blank>|\[BLANK\])\s*$', block.content, re.IGNORECASE))
                # Check if paragraph has text THEN ends with <blank> / [BLANK] / _____ (Option B: trailing blank)
                trailing_blank_match = re.match(r'^(.+?)\s*(?:<(?:blank|BLANK)>|\[(?:blank|BLANK)\]|_{3,})\s*$', block.content, re.DOTALL) if not ends_with_blank else None

                if ends_with_blank:
                    # Flush right standalone blank line with dynamic Tab Leader 4 (wdTabLeaderUnderscore = 4)
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(printable_width_cm), Alignment=2, Leader=4)
                    sel.Font.Name = self.font_name
                    sel.Font.Size = self.font_size
                    sel.Font.Bold = 0
                    sel.Font.Underline = 0
                    sel.TypeText("\t")
                    sel.TypeParagraph()
                    sel.ParagraphFormat.TabStops.ClearAll()
                elif trailing_blank_match:
                    # Option B: text before <blank> rendered inline, blank filled dynamically to right margin via Leader=4
                    text_part = trailing_blank_match.group(1)
                    from uln_parser import parse_inline_spans as _pis
                    text_spans = _pis(text_part)
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(printable_width_cm), Alignment=2, Leader=4)
                    self.write_inline_spans(sel, text_spans)
                    sel.TypeText("\t")
                    sel.TypeParagraph()
                    sel.ParagraphFormat.TabStops.ClearAll()
                else:
                    items = split_line_into_option_items(block.content)
                    if len(items) > 1:
                        num_cols = len(items)
                        self.setup_tab_stops(sel, num_cols, left_indent_cm=0.0, printable_width_cm=printable_width_cm)
                        
                        from uln_parser import parse_inline_spans
                        for idx_item, item in enumerate(items):
                            spans = parse_inline_spans(item.strip())
                            self.write_inline_spans(sel, spans)
                            if idx_item < len(items) - 1:
                                sel.TypeText("\t")
                        sel.TypeParagraph()
                        sel.ParagraphFormat.TabStops.ClearAll()
                    else:
                        self.write_inline_spans(sel, block.spans)
                        sel.TypeParagraph()

            elif tag in ["P1", "P2"]:
                items = split_line_into_option_items(block.content)
                starts_with_num = bool(re.match(r'^\s*\d+[\.\)]', block.content))

                # Question lines starting with a number (1. A. ... B. ...) MUST start flush at 0cm left margin!
                left_indent_cm = 0.0 if starts_with_num else (0.5 if tag == "P1" else 1.0)

                sel.ParagraphFormat.SpaceBefore = 4 if tag == "P1" else 3
                sel.ParagraphFormat.SpaceAfter = 3
                sel.ParagraphFormat.KeepWithNext = False
                sel.ParagraphFormat.Alignment = 0

                # Check if paragraph is ONLY a standalone blank line _____ or <blank> or [BLANK]
                ends_with_blank = bool(re.search(r'^\s*(?:_{3,}|<blank>|\[BLANK\])\s*$', block.content, re.IGNORECASE))
                # Check if paragraph has text THEN ends with <blank> / [BLANK] / _____ (Option B: trailing blank)
                trailing_blank_match = re.match(r'^(.+?)\s*(?:<(?:blank|BLANK)>|\[(?:blank|BLANK)\]|_{3,})\s*$', block.content, re.DOTALL) if not ends_with_blank else None

                if ends_with_blank:
                    sel.ParagraphFormat.LeftIndent = cm_to_pt(left_indent_cm)
                    sel.ParagraphFormat.FirstLineIndent = 0
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(printable_width_cm), Alignment=2, Leader=4)
                    sel.Font.Name = self.font_name
                    sel.Font.Size = self.font_size
                    sel.Font.Bold = 0
                    sel.Font.Underline = 0
                    sel.TypeText("\t")
                    sel.TypeParagraph()
                    sel.ParagraphFormat.TabStops.ClearAll()
                elif trailing_blank_match:
                    # Option B: text before <blank> rendered inline, blank filled dynamically to right margin via Leader=4
                    text_part = trailing_blank_match.group(1)
                    from uln_parser import parse_inline_spans as _pis
                    text_spans = _pis(text_part)
                    sel.ParagraphFormat.LeftIndent = cm_to_pt(left_indent_cm)
                    sel.ParagraphFormat.FirstLineIndent = 0
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(printable_width_cm), Alignment=2, Leader=4)
                    self.write_inline_spans(sel, text_spans)
                    sel.TypeText("\t")
                    sel.TypeParagraph()
                    sel.ParagraphFormat.TabStops.ClearAll()
                else:
                    # Detect horizontal picture choice grids
                    pic_matches = list(re.finditer(r'(?:\d+\.\s*)?\[PIC:[^\]]+\]\s*_{2,}', block.content))
                    if len(pic_matches) >= 2:
                        num_cols = min(4, len(pic_matches))
                        self.setup_tab_stops(sel, num_cols, left_indent_cm=left_indent_cm, printable_width_cm=printable_width_cm)
                        
                        from uln_parser import parse_inline_spans
                        items = split_line_into_option_items(block.content)
                        for idx_item, item in enumerate(items):
                            spans = parse_inline_spans(item.strip())
                            self.write_inline_spans(sel, spans)
                            if (idx_item + 1) % num_cols == 0 or idx_item == len(items) - 1:
                                sel.TypeParagraph()
                            else:
                                sel.TypeText("\t")
                        sel.ParagraphFormat.TabStops.ClearAll()
                    else:
                        if len(items) > 1:
                            num_cols = len(items)
                            self.setup_tab_stops(sel, num_cols, left_indent_cm=left_indent_cm, printable_width_cm=printable_width_cm)

                            from uln_parser import parse_inline_spans
                            for idx_item, item in enumerate(items):
                                spans = parse_inline_spans(item.strip())
                                self.write_inline_spans(sel, spans)
                                if idx_item < len(items) - 1:
                                    sel.TypeText("\t")
                            sel.TypeParagraph()
                            sel.ParagraphFormat.TabStops.ClearAll()
                        else:
                            sel.ParagraphFormat.LeftIndent = cm_to_pt(left_indent_cm)
                            sel.ParagraphFormat.FirstLineIndent = 0
                            self.write_inline_spans(sel, block.spans)
                            sel.TypeParagraph()

            elif tag == "TAB2":
                # Find the TRUE start of this consecutive TAB2 run (scan backward)
                group_start = idx_block
                while group_start > 0 and blocks[group_start - 1].tag == "TAB2":
                    group_start -= 1

                # Collect the full group from start to end
                tab2_group = []
                lookahead = group_start
                while lookahead < len(blocks) and blocks[lookahead].tag == "TAB2":
                    tab2_group.append(blocks[lookahead])
                    lookahead += 1

                max_c1_len = max(len(b.col1) for b in tab2_group) if tab2_group else 10
                max_c2_len = max(len(b.col2) for b in tab2_group) if tab2_group else 10

                group_first_c1 = tab2_group[0].col1 if tab2_group else block.col1
                base_indent_cm = 0.5 if "P1" in group_first_c1 else (1.0 if "P2" in group_first_c1 else 0.0)

                # Algorithm: Calculate TAB2 start so longest Col 2 line touches printable_width_cm (right border)
                min_col2_start_cm = base_indent_cm + max(2.5, (max_c1_len * 0.16) + 0.6)
                c2_width_cm = (max_c2_len * 0.175) + 0.6
                ideal_col2_start_cm = printable_width_cm - c2_width_cm

                if ideal_col2_start_cm >= min_col2_start_cm:
                    col2_tab_pos_cm = ideal_col2_start_cm
                else:
                    col2_tab_pos_cm = min_col2_start_cm

                col1_needed_cm = col2_tab_pos_cm - base_indent_cm

                sel.ParagraphFormat.SpaceBefore = 3
                sel.ParagraphFormat.SpaceAfter = 3

                # Check if Column 2 is an answer blank (e.g. ______ or <blank> or [BLANK])
                col2_is_blank = bool(re.match(r'^\s*(?:Answer:\s*)?(?:_{2,}|<blank>|\[BLANK\])\s*$', block.col2, re.IGNORECASE))

                # Detect header row in TAB2 (e.g. A | B)
                is_header_row = (block == tab2_group[0] and len(block.col1.strip()) <= 10 and len(block.col2.strip()) <= 10 and not re.search(r'\d', block.col1))

                if is_header_row:
                    col1_center_cm = base_indent_cm + (col1_needed_cm / 2.0)
                    col2_center_cm = col2_tab_pos_cm + ((printable_width_cm - col2_tab_pos_cm) / 2.0)

                    sel.ParagraphFormat.LeftIndent = 0
                    sel.ParagraphFormat.FirstLineIndent = 0
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(col1_center_cm), Alignment=1)  # wdAlignTabCenter = 1
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(col2_center_cm), Alignment=1)  # wdAlignTabCenter = 1

                    sel.TypeText("\t")
                    self.write_inline_spans(sel, block.col1_spans, default_bold=True)
                    sel.TypeText("\t")
                    self.write_inline_spans(sel, block.col2_spans, default_bold=True)
                    sel.TypeParagraph()

                elif col2_is_blank:
                    # FLUSH RIGHT TAB STOP at right margin (Alignment = 2, Leader = 4) for answer blanks
                    sel.ParagraphFormat.LeftIndent = cm_to_pt(base_indent_cm)
                    sel.ParagraphFormat.FirstLineIndent = 0
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(printable_width_cm), Alignment=2, Leader=4)  # wdAlignTabRight = 2, wdTabLeaderUnderscore = 4

                    self.write_inline_spans(sel, block.col1_spans)
                    sel.TypeText("\t")
                    sel.TypeParagraph()

                else:
                    col2_trim = block.col2.strip()
                    m_opt = re.match(r'^\s*(?:(?:\*\*|\*|\[|\(?)*([a-zA-Z])[\.\)](?:\*\*|\*|\]|\}|\{u\}|\))*)\s+(.*)$', col2_trim)

                    # Single tab stop: col1 text → \t → col2 content (option letter + body inline)
                    sel.ParagraphFormat.LeftIndent = cm_to_pt(col2_tab_pos_cm)
                    sel.ParagraphFormat.FirstLineIndent = cm_to_pt(-col1_needed_cm)
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(col2_tab_pos_cm), Alignment=0)

                    self.write_inline_spans(sel, block.col1_spans)
                    sel.TypeText("\t")

                    if m_opt and not block.pic:
                        from uln_parser import parse_inline_spans as _pis
                        col2_formatted = f"**{m_opt.group(1).upper()}.** {m_opt.group(2).strip()}"
                        self.write_inline_spans(sel, _pis(col2_formatted))
                    elif block.pic:
                        self.render_pic(sel, doc, block.pic)
                    else:
                        self.write_inline_spans(sel, block.col2_spans)

                    sel.TypeParagraph()

                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.FirstLineIndent = 0
                sel.ParagraphFormat.TabStops.ClearAll()

            elif tag == "TABLE":
                if block.table_data:
                    self.render_table(sel, doc, block.table_data, printable_width_cm)

            elif tag == "PIC_GRID":
                self.render_pic_grid(sel, doc, block.children, printable_width_cm)

            elif tag == "BOX":
                self.render_box_shape(sel, doc, word, block, printable_width_cm)

            elif tag == "QUOTE":
                # Reading passage: standard body text (Left/Right Indent = 0), only first line indented (0.75 cm), justified
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.RightIndent = 0
                sel.ParagraphFormat.FirstLineIndent = cm_to_pt(0.75)
                sel.ParagraphFormat.SpaceBefore = 0
                sel.ParagraphFormat.SpaceAfter = 4
                sel.ParagraphFormat.Alignment = 3  # wdAlignParagraphJustify = 3
                self.write_inline_spans(sel, block.spans, default_italic=False)
                sel.TypeParagraph()
                sel.ParagraphFormat.RightIndent = 0
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.FirstLineIndent = 0

            elif tag == "PIC":
                if block.pic:
                    sel.ParagraphFormat.LeftIndent = 0
                    sel.ParagraphFormat.FirstLineIndent = 0
                    sel.ParagraphFormat.SpaceBefore = 6
                    sel.ParagraphFormat.SpaceAfter = 6
                    if block.pic.pos == "center":
                        sel.ParagraphFormat.Alignment = 1  # Center
                    else:
                        sel.ParagraphFormat.Alignment = 0  # Left

                    self.render_pic(sel, doc, block.pic)
                    sel.TypeParagraph()

            try:
                word.ActiveWindow.ScrollIntoView(sel.Range, True)
            except Exception:
                try:
                    sel.ScrollIntoView()
                except Exception:
                    pass

            idx_block += 1

    def render_pic(self, sel, doc, pic: PicInfo):
        """Renders an image file from user queue in order, or falls back to 'test pic/' folder."""
        target_path = self.get_next_image_path(pic)

        if target_path and os.path.exists(target_path):
            try:
                shape = sel.InlineShapes.AddPicture(FileName=os.path.abspath(target_path))
                if pic.size == "small":
                    shape.Width = cm_to_pt(3.5)
                    shape.Height = cm_to_pt(2.5)
                elif pic.size == "large":
                    shape.Width = cm_to_pt(10.0)
                    shape.Height = cm_to_pt(6.5)
                else:
                    shape.Width = cm_to_pt(6.0)
                    shape.Height = cm_to_pt(4.0)
                return
            except Exception as e:
                print(f"[ULNRenderer] Warning adding picture {target_path}: {e}")

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
        try:
            sel.Font.ColorIndex = 0
        except Exception:
            pass

    def render_box_shape(self, sel, doc, word, block: ULNBlock, printable_width_cm: float):
        """
        Renders Word Bank / Callout Box using MS Word Rounded Rectangle Shape (msoShapeRoundedRectangle = 5)
        with text inserted directly inside shape.TextFrame with 0mm margins (0.0pt) on all sides.
        """
        printable_width_pt = cm_to_pt(printable_width_cm)
        
        if '|' in block.content:
            words = [w.strip() for w in block.content.split('|') if w.strip()]
        else:
            words = [w.strip() for w in block.content.split() if w.strip()]

        if not words:
            return

        N = len(words)
        if N <= 5:
            cols = N
        elif N <= 9:
            cols = 4
        else:
            cols = 5

        max_len_all = max(len(w) for w in words)
        char_w_pt = 6.0  # Average 12pt character width
        col_width_pt = (max_len_all * char_w_pt) + 16.0

        last_col_words = [words[i] for i in range(cols - 1, N, cols)] if N >= cols else [words[-1]]
        last_col_max_len = max(len(w) for w in last_col_words) if last_col_words else 8
        last_col_text_w_pt = last_col_max_len * char_w_pt

        text_group_width_pt = min(printable_width_pt, ((cols - 1) * col_width_pt) + last_col_text_w_pt)
        left_offset_pt = max(0.0, (printable_width_pt - text_group_width_pt) / 2.0)

        num_rows = (N + cols - 1) // cols
        box_width_pt = text_group_width_pt + 16.0
        box_height_pt = (num_rows * 18.0) + 8.0

        anchor_range = sel.Range.Duplicate

        try:
            shape = doc.Shapes.AddShape(
                5,  # msoShapeRoundedRectangle = 5
                0,
                0,
                box_width_pt,
                box_height_pt,
                Anchor=anchor_range
            )
            shape.RelativeHorizontalPosition = 0  # wdRelativeHorizontalPositionMargin = 0
            shape.RelativeVerticalPosition = 2    # wdRelativeVerticalPositionParagraph = 2
            shape.Left = left_offset_pt - 8.0
            shape.Top = 0.0

            # Set text box margins of shape to all 0mm (0.0pt) left, right, top, bottom
            tf = shape.TextFrame
            tf.MarginLeft = 0.0
            tf.MarginRight = 0.0
            tf.MarginTop = 0.0
            tf.MarginBottom = 0.0

            shape.Fill.Visible = False  # Transparent fill
            shape.Line.Weight = 1.0     # 1pt rounded border line
            shape.Line.ForeColor.RGB = 0  # Black border line

            # Select inside shape text frame to write text directly into shape
            tf.TextRange.Select()
            shape_sel = word.Selection

            shape_sel.ParagraphFormat.SpaceBefore = 4
            shape_sel.ParagraphFormat.SpaceAfter = 4
            shape_sel.ParagraphFormat.Alignment = 0  # Left align inside tab stops
            shape_sel.ParagraphFormat.LeftIndent = 8.0

            shape_sel.ParagraphFormat.TabStops.ClearAll()
            for c in range(1, cols):
                shape_sel.ParagraphFormat.TabStops.Add(Position=8.0 + (col_width_pt * c), Alignment=0)

            lines = [words[i:i + cols] for i in range(0, N, cols)]
            for idx_line, chunk in enumerate(lines):
                for idx_w, word_txt in enumerate(chunk):
                    shape_sel.Font.Name = self.font_name
                    shape_sel.Font.Size = self.font_size
                    shape_sel.Font.Bold = 1
                    shape_sel.TypeText(word_txt)
                    if idx_w < len(chunk) - 1:
                        shape_sel.Font.Bold = 0
                        shape_sel.TypeText("\t")
                if idx_line < len(lines) - 1:
                    shape_sel.TypeParagraph()

            # Return selection back to main document body after shape
            doc.Range(anchor_range.End, anchor_range.End).Select()
            sel = word.Selection
            sel.TypeParagraph()

        except Exception as e:
            print(f"[ULNRenderer] Warning creating shape with text frame: {e}")

    def render_table(self, sel, doc, tdata, printable_width_cm: float):
        """
        Renders [TABLE] block structure:
        - If borderless: uses divided paragraph tab stop columns with custom spacing.
        - If bordered: uses a native MS Word table with cell margins & 1.0pt gridlines.
        - Empty/blank cells (______): ignores literal underscore text inside cell bodies so bottom border acts as answer line!
        """
        if not tdata.rows:
            return

        num_rows = len(tdata.rows)
        num_cols = max(len(r.cells) for r in tdata.rows)

        if tdata.borderless:
            # Check if any cell in borderless table contains a [PIC] tag
            has_pic = any(
                "[PIC" in cell.content.upper() or parse_pic_tag(cell.content) is not None
                for row in tdata.rows for cell in row.cells
            )

            pic_info_found = None
            text_rows = []

            for row in tdata.rows:
                row_text_parts = []
                for cell in row.cells:
                    if "[PIC" in cell.content.upper() or parse_pic_tag(cell.content) is not None:
                        if not pic_info_found:
                            pic_info_found = parse_pic_tag(cell.content) or PicInfo(description="Activity Picture", pos="center", size="medium")
                    else:
                        if cell.content.strip() and not re.match(r'^\s*(?:_{2,}|<blank>|\[BLANK\])\s*$', cell.content, re.IGNORECASE):
                            row_text_parts.append(cell.content.strip())
                if row_text_parts:
                    text_rows.append(" ".join(row_text_parts))

            options_anchor_range = None
            opt_start_top_pt = 0.0

            # Render text rows as standard paragraphs (NO MS Word Table object!)
            for idx_r, txt_line in enumerate(text_rows):
                from uln_parser import parse_inline_spans
                spans = parse_inline_spans(txt_line)
                
                is_opt_line = bool(re.match(r'^\s*\*?\*?[A-Da-d][\.\)]', txt_line))
                left_ind_cm = 0.5 if is_opt_line else 0.0

                sel.ParagraphFormat.LeftIndent = cm_to_pt(left_ind_cm)
                sel.ParagraphFormat.FirstLineIndent = 0
                sel.ParagraphFormat.SpaceBefore = 3
                sel.ParagraphFormat.SpaceAfter = 3
                sel.ParagraphFormat.LineSpacingRule = 0  # Single Line Spacing
                sel.ParagraphFormat.Alignment = 0  # Left

                if idx_r == 1 or (is_opt_line and options_anchor_range is None):
                    # Capture anchor range and vertical position at top of Option A
                    try:
                        options_anchor_range = sel.Range.Duplicate
                        opt_start_top_pt = options_anchor_range.Information(6)  # wdVerticalPositionRelativeToPage = 6
                    except Exception:
                        pass

                self.write_inline_spans(sel, spans)
                sel.TypeParagraph()

            sel.ParagraphFormat.LeftIndent = 0

            # Measure vertical position at bottom of Option D
            opt_end_top_pt = 0.0
            try:
                opt_end_top_pt = sel.Range.Information(6)
            except Exception:
                pass

            calc_options_height_pt = max(cm_to_pt(2.5), opt_end_top_pt - opt_start_top_pt) if (opt_end_top_pt > opt_start_top_pt > 0) else cm_to_pt(2.8)

            if options_anchor_range is None:
                options_anchor_range = sel.Range.Duplicate

            # Floating Picture Placement (In Front of Text, Aligned Top Option A to Bottom Option D, Flush Right Margin)
            if pic_info_found or has_pic:
                target_path = self.get_next_image_path(pic_info_found)

                if target_path and os.path.exists(target_path):
                    try:
                        inline_shape = options_anchor_range.InlineShapes.AddPicture(FileName=os.path.abspath(target_path))
                        shape = inline_shape.ConvertToShape()

                        shape.WrapFormat.Type = 3  # msoWrapFront = 3 (In Front of Text)
                        shape.RelativeHorizontalPosition = 0  # wdRelativeHorizontalPositionMargin = 0
                        shape.RelativeVerticalPosition = 2    # wdRelativeVerticalPositionParagraph = 2

                        # Set height to match exact distance between Option A top and Option D bottom
                        shape.LockAspectRatio = 0  # Allow exact vertical fit
                        shape.Height = calc_options_height_pt
                        shape.Width = calc_options_height_pt * 1.25  # Maintain 5:4 aspect ratio

                        # Position flush right against page margin (printable_width_cm - shape_width_cm)
                        shape.Left = cm_to_pt(printable_width_cm) - shape.Width
                        shape.Top = cm_to_pt(0.0)  # Aligned with top of Option A
                    except Exception as pic_err:
                        print(f"[ULNRenderer] Floating image positioning error: {pic_err}")

            else:
                # Tab stop fallback for simple text-only borderless tables
                sel.ParagraphFormat.SpaceBefore = 3
                sel.ParagraphFormat.SpaceAfter = 3
                self.setup_tab_stops(sel, num_cols, left_indent_cm=0.5, printable_width_cm=printable_width_cm)

                for row in tdata.rows:
                    for idx_c, cell in enumerate(row.cells):
                        if not re.match(r'^\s*(?:_{2,}|<blank>|\[BLANK\])\s*$', cell.content, re.IGNORECASE):
                            self.write_inline_spans(sel, cell.spans, default_bold=cell.is_header)
                        else:
                            sel.TypeText("___________")
                        if idx_c < len(row.cells) - 1:
                            sel.TypeText("\t")
                    sel.TypeParagraph()

                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.TabStops.ClearAll()

        else:
            table = doc.Tables.Add(Range=sel.Range, NumRows=num_rows, NumColumns=num_cols)
            try:
                # wdAutoFitContent = 1: columns auto-fit to content width, prevents text wrapping
                table.AutoFitBehavior(1)
            except Exception:
                pass
            try:
                table.Rows.Alignment = 1  # Center table on page
            except Exception:
                pass

            try:
                table.TopPadding = cm_to_pt(0.15)
                table.BottomPadding = cm_to_pt(0.15)
                table.LeftPadding = cm_to_pt(0.25)
                table.RightPadding = cm_to_pt(0.25)
            except Exception:
                pass

            # Apply 1.0pt single black border lines
            for border_id in [-1, -2, -3, -4, -5, -6]:  # Top, Left, Bottom, Right, InsideH, InsideV
                try:
                    table.Borders(border_id).LineStyle = 1
                    table.Borders(border_id).LineWidth = 8  # 1pt
                    table.Borders(border_id).Color = 0  # Black
                except Exception:
                    pass

            for r_idx, row in enumerate(tdata.rows):
                for c_idx, cell in enumerate(row.cells):
                    if c_idx < num_cols:
                        cell_obj = table.Cell(r_idx + 1, c_idx + 1)
                        cell_obj.VerticalAlignment = 1  # wdCellVerticalAlignmentCenter = 1
                        
                        p_range = cell_obj.Range
                        p_range.ParagraphFormat.SpaceBefore = 2
                        p_range.ParagraphFormat.SpaceAfter = 2
                        p_range.ParagraphFormat.Alignment = 1 if cell.is_header else 0

                        sel.Start = cell_obj.Range.Start

                        # MANDATE: If cell content is _____, <blank>, [BLANK], or empty, IGNORE literal text so bottom border acts as answer line!
                        is_blank_cell = bool(re.match(r'^\s*(?:_{2,}|<blank>|\[BLANK\])\s*$', cell.content, re.IGNORECASE))
                        if not is_blank_cell and cell.content.strip():
                            self.write_inline_spans(sel, cell.spans, default_bold=cell.is_header)

            sel.Start = table.Range.End

    def render_pic_grid(self, sel, doc, children: List[ULNBlock], printable_width_cm: float):
        """
        Renders [PIC_GRID] using divided paragraph tab stops (NO MS Word Table object!).
        Row 1: 4 pictures placed across 4 tab stops (Single line spacing).
        Row 2: 4 captions (<number>. _________) placed across 4 tab stops matching picture width.
        """
        if not children:
            return

        cols = 4
        col_width_cm = printable_width_cm / cols

        # Process in chunks of 4 items per row
        for i in range(0, len(children), cols):
            chunk = children[i:i + cols]

            # -------------------------------------------------------------
            # LINE 1: PICTURES LINE (4 tab-aligned pictures, Single Line Spacing)
            # -------------------------------------------------------------
            sel.ParagraphFormat.LeftIndent = 0
            sel.ParagraphFormat.FirstLineIndent = 0
            sel.ParagraphFormat.LineSpacingRule = 0  # Single Line Spacing
            sel.ParagraphFormat.SpaceBefore = 4
            sel.ParagraphFormat.SpaceAfter = 2
            sel.ParagraphFormat.TabStops.ClearAll()

            for c in range(1, len(chunk)):
                tab_pos_cm = col_width_cm * c
                sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(tab_pos_cm), Alignment=0)

            for idx_c, item in enumerate(chunk):
                # Auto-determine picture dimensions by code (small 3.5cm x 2.5cm)
                pic_info = item.pic or PicInfo(description=f"Activity Image {i + idx_c + 1}", size="small")
                pic_info.size = "small"
                self.render_pic(sel, doc, pic_info)

                if idx_c < len(chunk) - 1:
                    sel.TypeText("\t")

            sel.TypeParagraph()
            sel.ParagraphFormat.TabStops.ClearAll()

            # -------------------------------------------------------------
            # LINE 2: CAPTIONS LINE (<number>. _________)
            # -------------------------------------------------------------
            sel.ParagraphFormat.LeftIndent = 0
            sel.ParagraphFormat.FirstLineIndent = 0
            sel.ParagraphFormat.SpaceBefore = 2
            sel.ParagraphFormat.SpaceAfter = 8
            sel.ParagraphFormat.TabStops.ClearAll()

            for c in range(1, len(chunk)):
                tab_pos_cm = col_width_cm * c
                sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(tab_pos_cm), Alignment=0)

            for idx_c, item in enumerate(chunk):
                item_idx = i + idx_c + 1
                num_match = re.match(r'^\s*(\d+[\.\)])\s*', item.content)
                num_str = f"{num_match.group(1)} " if num_match else f"{item_idx}. "

                sel.Font.Name = self.font_name
                sel.Font.Size = self.font_size
                sel.Font.Bold = 1
                sel.TypeText(num_str)

                sel.Font.Bold = 0
                sel.Font.Underline = 0
                sel.TypeText("_________")

                if idx_c < len(chunk) - 1:
                    sel.TypeText("\t")

            sel.TypeParagraph()
            sel.ParagraphFormat.TabStops.ClearAll()
