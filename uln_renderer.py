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
    hex_match = re.search(r'#([0-9a-fA-F]{6})', color_str)
    if hex_match:
        hex_val = hex_match.group(1).lower()
        r = int(hex_val[0:2], 16)
        g = int(hex_val[2:4], 16)
        b = int(hex_val[4:6], 16)
        return r + (g * 256) + (b * 65536)
    if clean in COLOR_NAME_TO_RGB:
        return COLOR_NAME_TO_RGB[clean]
    return None



def extract_question_prefix_and_body(text: str) -> tuple[Optional[str], Optional[str], Optional[str], str]:
    """
    Extracts (full_prefix, delim_char, q_num, body_text) when formatted with '#' placeholder:
    - #1. -> ('', '.', '1', 'Mike...') -> Word list: '1.'
    - Question #1. -> ('Question ', '.', '1', 'What...') -> Word list: 'Question 1.'
    - Question #1: -> ('Question ', ':', '1', 'What...') -> Word list: 'Question 1:'
    - Câu #1: -> ('Câu ', ':', '1', '...') -> Word list: 'Câu 1:'
    - Task #1. -> ('Task ', '.', '1', '...') -> Word list: 'Task 1.'
    Returns (full_prefix, delim_char, q_num, body_text) if matched, else (None, None, None, original_text).
    """
    if not text or '#' not in text:
        return None, None, None, text

    pattern = r'^\s*(?:\*\*)?([A-Za-zÀ-ỹ\s]+?)?#(\d+)([\.\)\:\-]|\s*:\s*|\s*\.\s*)?\s*(?:\*\*)?[:\.\)]?\s*(.*)$'
    m = re.match(pattern, text, re.IGNORECASE)
    if m:
        prefix_word = m.group(1).strip() if m.group(1) else ""
        q_num = m.group(2)
        delimiter = m.group(3).strip() if m.group(3) else "."
        body = m.group(4).strip()
        body = re.sub(r'^(?:\*\*|[:\.\-\)])\s*', '', body).strip()

        full_prefix = f"{prefix_word} " if prefix_word else ""
        delim_char = ":" if ":" in delimiter else ("." if "." in delimiter else (delimiter if delimiter else "."))
        return full_prefix, delim_char, q_num, body
    return None, None, None, text


def split_line_into_option_items(line_text: str) -> List[str]:
    """
    Splits line text into multi-column items automatically by detecting:
    1. Explicit \\t characters
    2. Multiple sequential option choices like A. text B. text C. text D. text or a. b. c. d. or 1. 2. 3. 4.
    """
    if not line_text:
        return []

    if '\t' in line_text:
        return [x.strip() for x in line_text.split('\t') if x.strip()]

    # Extract potential question number prefix (e.g. "#1. ", "Question #1: ")
    pref, delim, q_num, rest_text = extract_question_prefix_and_body(line_text)
    q_prefix = f"{pref}#{q_num}{delim} " if q_num else ""

    # Search for option choices like A. B. C. D. or a. b. c. d. or (A) (B) (C) or 1. 2. 3. 4.

    pattern = r'(?:^|\s+)(?:(?:\*\*|\*|\[|\(?)*([a-zA-Z0-9][\.\)])(?:\*\*|\*|\]|\}|\{u\}|\))*)(?=\s+|$)'
    matches = list(re.finditer(pattern, rest_text))
    if len(matches) >= 2:
        labels = [re.sub(r'[\*\_\`\[\]\(\)\.\s]', '', m.group(1)) for m in matches]
        
        is_seq = False
        if all(len(l) == 1 and l.isalpha() for l in labels):
            if labels[0] == 'A' and all(ord(l) == ord('A') + i for i, l in enumerate(labels)):
                is_seq = True
            elif labels[0] == 'a' and all(ord(l) == ord('a') + i for i, l in enumerate(labels)):
                is_seq = True
        elif all(l.isdigit() for l in labels):
            nums = [int(l) for l in labels]
            if nums[0] == 1 and all(n == 1 + i for i, n in enumerate(nums)):
                is_seq = True

        if is_seq:
            items = []
            for i in range(len(matches)):
                start = matches[i].start()
                end = matches[i + 1].start() if i + 1 < len(matches) else len(rest_text)
                item_txt = rest_text[start:end].strip()
                if i == 0 and q_prefix:
                    item_txt = f"{q_prefix.strip()} {item_txt}"
                items.append(item_txt)
            return items

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
        self.enable_page_numbers = self.settings.get("enable_page_numbers", True)
        # Question & Option Styling Settings
        self.question_prefix = self.settings.get("question_prefix", "")  # e.g. "", "Question ", "Câu ", "Task "
        self.question_delimiter = self.settings.get("question_delimiter", ".")  # e.g. ".", ":", ")", "-"
        self.question_color = self.settings.get("question_color", "#000000")  # Hex or color name
        self.opt_color = self.settings.get("opt_color", "#000000")  # Hex or color name


        self.user_images = list(self.settings.get("user_images", []))
        self.user_img_idx = 0
        self.is_first_question_in_num_block = False
        self.current_group_opt_cols = None
        self.current_group_max_item_len = None
        self.last_rendered_tag = None

    def get_effective_number_format(self, extracted_pref: Optional[str], extracted_delim: Optional[str]) -> str:
        """
        Determines the effective list NumberFormat string:
        - Ensures a proper single space between prefix word and %1 (e.g. 'Question ' -> 'Question %1.').
        - If delimiter was selected in GUI settings (and is non-default, e.g. ':', ')', '-'), use GUI delimiter.
          Otherwise use extracted delimiter from text or global default.
        """
        pref = extracted_pref if (extracted_pref and extracted_pref.strip()) else (self.question_prefix or "")
        if pref and pref.strip():
            pref = f"{pref.strip()} "
        else:
            pref = ""

        if self.question_delimiter and self.question_delimiter != ".":
            delim = self.question_delimiter
        else:
            delim = extracted_delim if (extracted_delim and extracted_delim.strip()) else (self.question_delimiter or ".")
        return f"{pref}%1{delim}"




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
                footer.Range.Font.Name = self.font_name
                footer.Range.Font.Size = 10
                footer.Range.Text = ""

                sel_range = footer.Range
                doc.Fields.Add(Range=sel_range, Type=-1, Text="PAGE")  # wdFieldPage = -1
                footer.Range.InsertAfter(" / ")
                end_range = footer.Range
                end_range.Collapse(0)  # wdCollapseEnd = 0
                doc.Fields.Add(Range=end_range, Type=-1, Text="NUMPAGES")

                # Center footer page numbers across all paragraphs in footer
                try:
                    footer.Range.ParagraphFormat.TabStops.ClearAll()
                    for p in footer.Range.Paragraphs:
                        p.Alignment = 1  # wdAlignParagraphCenter = 1
                    footer.Range.ParagraphFormat.Alignment = 1
                except Exception:
                    pass

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
                    sel.Font.Color = 0  # Pure Black RGB(0,0,0)
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

    def setup_tab_stops(self, sel, num_cols: int, left_indent_cm: float, printable_width_cm: float, max_item_len: int = 0) -> float:
        """Calculates exact equal column division for tab stops so option columns ALWAYS align vertically across all questions."""
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

                # Check if paragraph is ONLY a standalone blank line _____ or <blank> or [BLANK], optionally followed by symbol/punct
                blank_symbol_match = re.match(r'^\s*(?:_{3,}|<(?:blank|BLANK)>|\[(?:blank|BLANK)\])\s*([?\.\!:,;]?)\s*$', block.content, re.IGNORECASE)
                # Check if paragraph has text THEN ends with <blank> / [BLANK] / _____ (Option B: trailing blank), ONLY for sentence transformation arrows (→ / ->) or 15+ long underscores
                is_transform_or_long = bool(re.match(r'^\s*(?:→|->)', block.content)) or bool(re.search(r'_{15,}', block.content))
                trailing_blank_symbol_match = re.match(r'^(.+?)\s*(?:<(?:blank|BLANK)>|\[(?:blank|BLANK)\]|_{3,})\s*([?\.\!:,;]?)\s*$', block.content, re.DOTALL | re.IGNORECASE) if (not blank_symbol_match and is_transform_or_long) else None

                if blank_symbol_match:
                    trailing_sym = blank_symbol_match.group(1).strip()
                    if self.last_rendered_tag == "BOX":
                        sel.ParagraphFormat.SpaceBefore = 14
                    # Flush right standalone blank line with dynamic Tab Leader 4 (wdTabLeaderUnderscore = 4)
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(printable_width_cm), Alignment=2, Leader=4)
                    sel.Font.Name = self.font_name
                    sel.Font.Size = self.font_size
                    sel.Font.Bold = 0
                    sel.Font.Underline = 0
                    if trailing_sym:
                        sel.TypeText(f"\t{trailing_sym}")
                    else:
                        sel.TypeText("\t")
                    sel.TypeParagraph()
                    sel.ParagraphFormat.TabStops.ClearAll()
                elif trailing_blank_symbol_match:
                    # Option B: text before <blank> rendered inline, blank filled dynamically to right margin via Leader=4
                    text_part = trailing_blank_symbol_match.group(1)
                    trailing_sym = trailing_blank_symbol_match.group(2).strip()

                    q_num_match = re.match(r'^\s*(?:#?(\d+)[\.\)]|Question\s+#?(\d+)[\.\)]?|Câu\s+#?(\d+)[\.\)]?)\s*(.*)$', text_part, re.IGNORECASE)
                    if q_num_match and q_num_match.group(4).strip():
                        self.apply_native_numbered_list(word, sel)
                        if self.last_rendered_tag == "BOX":
                            sel.ParagraphFormat.SpaceBefore = 14
                        text_part = q_num_match.group(4).strip()
                    else:
                        try:
                            sel.Range.ListFormat.RemoveNumbers()
                        except Exception:
                            pass
                        if self.last_rendered_tag == "BOX":
                            sel.ParagraphFormat.SpaceBefore = 14

                    from uln_parser import parse_inline_spans as _pis
                    text_spans = _pis(text_part)
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(printable_width_cm), Alignment=2, Leader=4)
                    self.write_inline_spans(sel, text_spans)
                    if trailing_sym:
                        sel.TypeText(f"\t{trailing_sym}")
                    else:
                        sel.TypeText("\t")
                    sel.TypeParagraph()
                    sel.ParagraphFormat.TabStops.ClearAll()
                else:
                    pref, delim, q_num, body_text = extract_question_prefix_and_body(block.content)
                    if q_num is not None:
                        num_fmt = self.get_effective_number_format(pref, delim)
                        self.apply_native_numbered_list(word, sel, q_num=q_num, number_format=num_fmt)
                    else:
                        try:
                            sel.Range.ListFormat.RemoveNumbers()
                        except Exception:
                            pass

                    items = split_line_into_option_items(body_text)
                    if len(items) > 1:
                        num_cols = self.calculate_optimal_option_cols(items, 0.0, printable_width_cm)
                        max_item_len = max(len(i) for i in items)
                        self.setup_tab_stops(sel, num_cols, left_indent_cm=0.0, printable_width_cm=printable_width_cm, max_item_len=max_item_len)
                        
                        from uln_parser import parse_inline_spans
                        for idx_item, item in enumerate(items):
                            spans = parse_inline_spans(item.strip())
                            self.write_inline_spans(sel, spans)
                            if (idx_item + 1) % num_cols == 0 or idx_item == len(items) - 1:
                                sel.TypeParagraph()
                            else:
                                sel.TypeText("\t")
                        sel.ParagraphFormat.TabStops.ClearAll()
                    else:
                        if self.last_rendered_tag == "BOX":
                            sel.ParagraphFormat.SpaceBefore = 0
                        from uln_parser import parse_inline_spans
                        body_spans = parse_inline_spans(body_text)
                        self.write_inline_spans(sel, body_spans)
                        sel.TypeParagraph()


            elif tag in ["P1", "P2"]:
                pref, delim, q_num, content_to_render = extract_question_prefix_and_body(block.content)
                left_indent_cm = 0.0 if (q_num is not None) else (0.5 if tag == "P1" else 1.0)

                if q_num is not None:
                    num_fmt = self.get_effective_number_format(pref, delim)
                    self.apply_native_numbered_list(word, sel, q_num=q_num, number_format=num_fmt)
                else:
                    try:
                        sel.Range.ListFormat.RemoveNumbers()
                    except Exception:
                        pass



                items = split_line_into_option_items(content_to_render)


                space_before_p1 = 14 if (self.last_rendered_tag == "BOX") else (4 if tag == "P1" else 3)
                sel.ParagraphFormat.SpaceBefore = space_before_p1
                sel.ParagraphFormat.SpaceAfter = 3
                sel.ParagraphFormat.KeepWithNext = False
                sel.ParagraphFormat.Alignment = 0

                # Check if paragraph is ONLY a standalone blank line _____ or <blank> or [BLANK], optionally followed by symbol/punct
                blank_symbol_match = re.match(r'^\s*(?:_{3,}|<(?:blank|BLANK)>|\[(?:blank|BLANK)\])\s*([?\.\!:,;]?)\s*$', content_to_render, re.IGNORECASE)
                # Check if paragraph has text THEN ends with <blank> / [BLANK] / _____ (Option B: trailing blank), ONLY for sentence transformation arrows (→ / ->) or 15+ long underscores
                is_transform_or_long = bool(re.match(r'^\s*(?:→|->)', content_to_render)) or bool(re.search(r'_{15,}', content_to_render))
                trailing_blank_symbol_match = re.match(r'^(.+?)\s*(?:<(?:blank|BLANK)>|\[(?:blank|BLANK)\]|_{3,})\s*([?\.\!:,;]?)\s*$', content_to_render, re.DOTALL | re.IGNORECASE) if (not blank_symbol_match and is_transform_or_long) else None

                if blank_symbol_match:
                    trailing_sym = blank_symbol_match.group(1).strip()
                    sel.ParagraphFormat.LeftIndent = cm_to_pt(left_indent_cm)
                    sel.ParagraphFormat.FirstLineIndent = 0
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(printable_width_cm), Alignment=2, Leader=4)
                    sel.Font.Name = self.font_name
                    sel.Font.Size = self.font_size
                    sel.Font.Bold = 0
                    sel.Font.Underline = 0
                    if trailing_sym:
                        sel.TypeText(f"\t{trailing_sym}")
                    else:
                        sel.TypeText("\t")
                    sel.TypeParagraph()
                    sel.ParagraphFormat.TabStops.ClearAll()
                elif trailing_blank_symbol_match:
                    # Option B: text before <blank> rendered inline, blank filled dynamically to right margin via Leader=4
                    text_part = trailing_blank_symbol_match.group(1)
                    trailing_sym = trailing_blank_symbol_match.group(2).strip()
                    from uln_parser import parse_inline_spans as _pis
                    text_spans = _pis(text_part)
                    sel.ParagraphFormat.LeftIndent = cm_to_pt(left_indent_cm)
                    sel.ParagraphFormat.FirstLineIndent = 0
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(printable_width_cm), Alignment=2, Leader=4)
                    self.write_inline_spans(sel, text_spans)
                    if trailing_sym:
                        sel.TypeText(f"\t{trailing_sym}")
                    else:
                        sel.TypeText("\t")
                    sel.TypeParagraph()
                    sel.ParagraphFormat.TabStops.ClearAll()
                else:
                    # Detect horizontal picture choice grids
                    pic_matches = list(re.finditer(r'(?:\d+\.\s*)?\[PIC:[^\]]+\]\s*_{2,}', content_to_render))
                    if len(pic_matches) >= 2:
                        num_cols = min(4, len(pic_matches))
                        self.setup_tab_stops(sel, num_cols, left_indent_cm=left_indent_cm, printable_width_cm=printable_width_cm)
                        
                        from uln_parser import parse_inline_spans
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
                            num_cols = self.calculate_optimal_option_cols(items, left_indent_cm, printable_width_cm)
                            self.setup_tab_stops(sel, num_cols, left_indent_cm=left_indent_cm, printable_width_cm=printable_width_cm)

                            from uln_parser import parse_inline_spans
                            for idx_item, item in enumerate(items):
                                spans = parse_inline_spans(item.strip())
                                self.write_inline_spans(sel, spans)
                                if (idx_item + 1) % num_cols == 0 or idx_item == len(items) - 1:
                                    sel.TypeParagraph()
                                else:
                                    sel.TypeText("\t")
                            sel.ParagraphFormat.TabStops.ClearAll()
                        else:
                            sel.ParagraphFormat.LeftIndent = cm_to_pt(left_indent_cm)
                            sel.ParagraphFormat.FirstLineIndent = 0
                            from uln_parser import parse_inline_spans
                            c_spans = parse_inline_spans(content_to_render)
                            self.write_inline_spans(sel, c_spans)
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

                space_before_tab2 = 14 if (self.last_rendered_tag == "BOX") else 3
                sel.ParagraphFormat.SpaceBefore = space_before_tab2
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
                    # Dynamic 2-Column Error Correction Layout:
                    # Column 2 starts right after longest Column 1 sentence (max_c1_len) with max 3.5cm answer blank
                    col2_tab_pos_cm = base_indent_cm + (max_c1_len * 0.18) + 0.8
                    col2_tab_pos_cm = min(col2_tab_pos_cm, printable_width_cm - 1.5)
                    avail_w_cm = max(1.5, printable_width_cm - col2_tab_pos_cm)
                    blank_w_cm = min(3.5, avail_w_cm)
                    num_underscores = int(blank_w_cm * 5.5)

                    sel.ParagraphFormat.LeftIndent = cm_to_pt(base_indent_cm)
                    sel.ParagraphFormat.FirstLineIndent = 0
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(col2_tab_pos_cm), Alignment=0)

                    pref, delim, q_num, c1_body = extract_question_prefix_and_body(block.col1)
                    if q_num is not None and c1_body.strip():
                        num_fmt = self.get_effective_number_format(pref, delim)
                        self.apply_native_numbered_list(word, sel, q_num=q_num, number_format=num_fmt)
                        from uln_parser import parse_inline_spans
                        c1_spans = parse_inline_spans(c1_body.strip())
                        self.write_inline_spans(sel, c1_spans)

                    else:
                        try:
                            sel.Range.ListFormat.RemoveNumbers()
                        except Exception:
                            pass
                        self.write_inline_spans(sel, block.col1_spans)



                    sel.TypeText(f"\t{'_' * num_underscores}")
                    sel.TypeParagraph()

                else:
                    col2_trim = block.col2.strip()
                    m_opt = re.match(r'^\s*(?:(?:\*\*|\*|\[|\(?)*([a-zA-Z])[\.\)](?:\*\*|\*|\]|\}|\{u\}|\))*)\s+(.*)$', col2_trim)

                    # Single tab stop: col1 text → \t → col2 content (option letter + body inline)
                    sel.ParagraphFormat.LeftIndent = cm_to_pt(col2_tab_pos_cm)
                    sel.ParagraphFormat.FirstLineIndent = cm_to_pt(-col1_needed_cm)
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(col2_tab_pos_cm), Alignment=0)

                    # Format Column 1 with Question Numbering if present (e.g. #1. or Question #1.)
                    pref, delim, q_num, c1_body = extract_question_prefix_and_body(block.col1)
                    if q_num is not None:
                        eff_pref = pref if (pref and pref.strip()) else (self.question_prefix or "")
                        if eff_pref and eff_pref.strip():
                            eff_pref = f"{eff_pref.strip()} "
                        else:
                            eff_pref = ""
                        eff_delim = self.question_delimiter if (self.question_delimiter and self.question_delimiter != ".") else (delim or ".")
                        num_str = f"{eff_pref}{q_num}{eff_delim} "

                        sel.Font.Name = self.font_name
                        sel.Font.Size = self.font_size
                        sel.Font.Bold = 1
                        sel.Font.Italic = 0
                        sel.Font.Underline = 0
                        q_color_int = parse_color_to_rgb_int(self.question_color)
                        if q_color_int is not None:
                            sel.Font.Color = q_color_int
                        else:
                            sel.Font.Color = 0
                        sel.TypeText(num_str)
                        sel.Font.Bold = 0
                        sel.Font.Color = 0

                        from uln_parser import parse_inline_spans
                        c1_spans = parse_inline_spans(c1_body.strip())
                        self.write_inline_spans(sel, c1_spans)
                    else:
                        self.write_inline_spans(sel, block.col1_spans)

                    sel.TypeText("\t")

                    if m_opt and not block.pic:
                        opt_let = f"{m_opt.group(1).upper()}."
                        opt_body = m_opt.group(2).strip()

                        sel.Font.Name = self.font_name
                        sel.Font.Size = self.font_size
                        sel.Font.Bold = 1
                        sel.Font.Italic = 0
                        sel.Font.Underline = 0
                        opt_color_int = parse_color_to_rgb_int(self.opt_color)
                        if opt_color_int is not None:
                            sel.Font.Color = opt_color_int
                        else:
                            sel.Font.Color = 0
                        sel.TypeText(f"{opt_let} ")
                        sel.Font.Bold = 0
                        sel.Font.Color = 0

                        from uln_parser import parse_inline_spans as _pis
                        self.write_inline_spans(sel, _pis(opt_body))
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
                    self.render_table(sel, doc, block.table_data, printable_width_cm, idx_block=idx_block, blocks=blocks)

            elif tag == "PIC_GRID":
                self.render_pic_grid(sel, doc, block.children, printable_width_cm)

            elif tag == "BOX":
                self.render_box_shape(sel, doc, word, block, printable_width_cm)

            elif tag == "NUM":
                self.render_num_container(sel, doc, word, block, printable_width_cm)

            elif tag == "OPT":
                if self.current_group_opt_cols is None:
                    consecutive_opts = []
                    k = idx_block
                    while k < len(blocks) and blocks[k].tag == "OPT":
                        consecutive_opts.append(blocks[k])
                        k += 1
                    if consecutive_opts:
                        g_cols, g_len = self.compute_group_option_params(consecutive_opts, printable_width_cm)
                        self.current_group_opt_cols = g_cols
                        self.current_group_max_item_len = g_len

                self.render_opt(sel, doc, word, block, printable_width_cm)

                # Reset group state after last consecutive OPT block
                if (idx_block + 1 >= len(blocks) or blocks[idx_block + 1].tag != "OPT") and not self.is_first_question_in_num_block:
                    self.current_group_opt_cols = None
                    self.current_group_max_item_len = None

            elif tag == "QUOTE":
                # Reading passage: standard body text (Left/Right Indent = 0), only first line indented (0.75 cm), justified
                space_before_quote = 14 if (self.last_rendered_tag == "BOX") else 0
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.RightIndent = 0
                sel.ParagraphFormat.FirstLineIndent = cm_to_pt(0.75)
                sel.ParagraphFormat.SpaceBefore = space_before_quote
                sel.ParagraphFormat.SpaceAfter = 4
                sel.ParagraphFormat.Alignment = 3  # wdAlignParagraphJustify = 3
                self.write_inline_spans(sel, block.spans, default_italic=False)
                sel.TypeParagraph()
                sel.ParagraphFormat.RightIndent = 0
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.FirstLineIndent = 0

            elif tag == "PIC":
                if block.pic:
                    space_before_pic = 14 if (self.last_rendered_tag == "BOX") else 6
                    sel.ParagraphFormat.LeftIndent = 0
                    sel.ParagraphFormat.FirstLineIndent = 0
                    sel.ParagraphFormat.SpaceBefore = space_before_pic
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

            self.last_rendered_tag = tag
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

    def apply_native_numbered_list(self, word, sel, q_num: Optional[str] = None, number_format: Optional[str] = None):
        """
        Applies native MS Word Numbered List:
        - List number is ALWAYS BOLD by default
        - Text and numbering are separated by a SINGLE SPACE (TrailingCharacter = wdTrailingSpace = 2), NOT a tab!
        - LeftIndent and FirstLineIndent are flush at 0.0 cm (left page border)
        - Starts a new independent list instance (ContinuePreviousList=False) on first question (q_num == '1' or is_first_question_in_num_block),
          and continues sequential list incrementing (ContinuePreviousList=True) on subsequent questions.
        - Supports custom prefix formats (e.g. 'Question %1.', 'Question %1:', 'Câu %1:', '%1.').
        """
        restart = self.is_first_question_in_num_block or (q_num == "1")
        self.is_first_question_in_num_block = False
        try:
            list_tpl = word.ListGalleries(2).ListTemplates(1)  # wdNumberGallery = 2 (Numbered List 1., 2., 3.)
            lvl = list_tpl.ListLevels(1)
            lvl.TrailingCharacter = 1  # wdTrailingSpace = 1 (SPACE separator, "Follow number with: Space")
            lvl.Font.Bold = 1          # ALWAYS BOLD number
            lvl.NumberPosition = 0
            lvl.TextPosition = 0
            lvl.NumberFormat = number_format if number_format else "%1."
            q_color_int = parse_color_to_rgb_int(self.question_color)
            if q_color_int is not None:
                lvl.Font.Color = q_color_int
            else:
                lvl.Font.Color = 0

            sel.Range.ListFormat.ApplyListTemplate(list_tpl, ContinuePreviousList=not restart)
            sel.ParagraphFormat.LeftIndent = 0
            sel.ParagraphFormat.FirstLineIndent = 0
            sel.Font.Bold = 0
            sel.Font.Color = 0  # Black body text

        except Exception:
            try:
                sel.Range.ListFormat.ApplyNumberDefault()
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.FirstLineIndent = 0
                sel.Font.Bold = 0
            except Exception:
                pass




    def clean_num_placeholders(self, b: ULNBlock):
        r"""Recursively cleans #(\d+) placeholders from content, columns, spans, tables, and child blocks."""
        def process_text_num(text: str) -> str:
            if not text:
                return text
            # Format #N to N so q_num_match can extract question number prefix
            return re.sub(r'#(\d+)', r'\1', text)

        if b.content:
            b.content = process_text_num(b.content)
        if b.col1:
            b.col1 = process_text_num(b.col1)
        if b.col2:
            b.col2 = process_text_num(b.col2)
        if b.spans:
            for span in b.spans:
                span.text = process_text_num(span.text)
        if b.col1_spans:
            for span in b.col1_spans:
                span.text = process_text_num(span.text)
        if b.col2_spans:
            for span in b.col2_spans:
                span.text = process_text_num(span.text)

        if b.table_data and b.table_data.rows:
            for row in b.table_data.rows:
                for cell in row.cells:
                    if cell.content:
                        cell.content = process_text_num(cell.content)
                    if cell.spans:
                        for span in cell.spans:
                            span.text = process_text_num(span.text)

        if b.children:
            for child in b.children:
                self.clean_num_placeholders(child)

    def compute_group_option_params(self, opt_blocks: List[ULNBlock], printable_width_cm: float):
        """Pre-scans all OPT blocks in an exercise group to compute uniform column count and tab stops."""
        all_items = []
        for b in opt_blocks:
            raw_text = b.content.strip()
            if '|' in raw_text:
                items = [x.strip() for x in raw_text.split('|') if x.strip()]
            elif '\n' in raw_text:
                items = [x.strip() for x in raw_text.split('\n') if x.strip()]
            else:
                items = split_line_into_option_items(raw_text)

            for idx_i, item in enumerate(items):
                pref, delim, q_num, clean_item = extract_question_prefix_and_body(item) if idx_i == 0 else (None, None, None, item)
                m_let = re.match(r'^\s*(?:(?:\*\*|\*|\[|\(?)*([a-zA-Z][\.\)])(?:\*\*|\*|\]|\}|\{u\}|\))*)\s*(.*)$', clean_item)
                item_str = f"{m_let.group(1)} {m_let.group(2)}" if m_let else f"A. {clean_item}"
                all_items.append(item_str)


        if all_items:
            left_indent_cm = 0.5
            cols = self.calculate_optimal_option_cols(all_items, left_indent_cm, printable_width_cm)
            max_len = max(len(i) for i in all_items)
            return cols, max_len
        return None, None

    def render_num_container(self, sel, doc, word, block: ULNBlock, printable_width_cm: float):
        """
        Renders auto-numbered container [NUM] ... [/NUM].
        Flags the first question in this section to start a new independent list.
        Pre-computes uniform option alignment across all questions in the exercise.
        """
        if not block.children:
            return

        # Flag that the first question in this NUM section starts a new list sequence at 1.
        self.is_first_question_in_num_block = True


        opt_blocks = [c for c in block.children if c.tag == "OPT"]
        old_cols, old_len = self.current_group_opt_cols, self.current_group_max_item_len
        if opt_blocks:
            g_cols, g_len = self.compute_group_option_params(opt_blocks, printable_width_cm)
            self.current_group_opt_cols = g_cols
            self.current_group_max_item_len = g_len

        try:
            self.render(block.children, doc, word)
        finally:
            self.current_group_opt_cols, self.current_group_max_item_len = old_cols, old_len

    def calculate_optimal_option_cols(self, items: List[str], left_indent_cm: float, printable_width_cm: float) -> int:
        """
        Calculates optimal column count (1, 2, 3, or 4 columns) for multiple-choice options
        so text wrapping NEVER occurs across available printable width.
        """
        N = len(items)
        if N <= 1:
            return 1

        remaining_width_cm = max(5.0, printable_width_cm - left_indent_cm)
        max_len = max(len(item) for item in items) if items else 0
        
        # Estimated option width in cm (0.165cm per char for 12pt Times New Roman + 0.4cm safety buffer for option letter)
        est_item_w_cm = (max_len * 0.165) + 0.4

        if N >= 4:
            # Standard 4-choice options <= 24 chars easily fit in 4 columns on 1 line across 16cm printable width
            if max_len <= 24 or (est_item_w_cm * 4) <= (remaining_width_cm + 0.5):
                return 4
            elif (est_item_w_cm * 2) <= (remaining_width_cm + 0.5):
                return 2
            else:
                return 1
        elif N == 3:
            if (est_item_w_cm * 3) <= remaining_width_cm:
                return 3
            else:
                return 1
        elif N == 2:
            if (est_item_w_cm * 2) <= remaining_width_cm:
                return 2
            else:
                return 1

        return 1

    def render_opt(self, sel, doc, word, block: ULNBlock, printable_width_cm: float):
        """
        Renders dedicated multiple-choice option container [OPT] ... [/OPT].
        Automatically formats option letters (A., B., C., D.) as bold, and calculates optimal column count
        (1, 2, 3, or 4 columns) based on max item length so text wrapping NEVER occurs.
        Handles questions with only options (Pronunciation/Stress/Odd-One-Out) keeping question number and options on 1 line.
        """
        raw_text = block.content.strip()
        if not raw_text:
            return

        # Split items by pipe '|' or line breaks if present
        if '|' in raw_text:
            raw_items = [x.strip() for x in raw_text.split('|') if x.strip()]
        elif '\n' in raw_text:
            raw_items = [x.strip() for x in raw_text.split('\n') if x.strip()]
        else:
            raw_items = split_line_into_option_items(raw_text)

        if not raw_items:
            return

        # Extract question number prefix if present in first item (e.g. #1., #1, 1., Question #1)
        q_num = None
        num_fmt = None
        if raw_items:
            pref, delim, q_num, first_body = extract_question_prefix_and_body(raw_items[0])
            if q_num is not None:
                raw_items[0] = first_body
                num_fmt = self.get_effective_number_format(pref, delim)

        formatted_items = []
        for idx, item in enumerate(raw_items):
            m = re.match(r'^\s*(?:(?:\*\*|\*|\[|\(?)*([a-zA-Z][\.\)])(?:\*\*|\*|\]|\}|\{u\}|\))*)\s*(.*)$', item)
            if m:
                opt_let = m.group(1).upper().rstrip('.')
                opt_letter = f"{opt_let}."
                body = m.group(2).strip()
            else:
                letter_char = chr(65 + idx)
                opt_letter = f"{letter_char}."
                body = item.strip()

            formatted_items.append((opt_letter, body))

        N = len(formatted_items)
        left_indent_cm = 0.0 if (q_num is not None) else 0.5
        
        # Use uniform group option parameters if set across this exercise
        if self.current_group_opt_cols is not None:
            cols = self.current_group_opt_cols
            max_item_len = self.current_group_max_item_len
        else:
            items_for_calc = [f"{let} {b}" for let, b in formatted_items]
            cols = self.calculate_optimal_option_cols(items_for_calc, left_indent_cm, printable_width_cm)
            max_item_len = max(len(i) for i in items_for_calc) if items_for_calc else 0

        sel.ParagraphFormat.SpaceBefore = 4
        sel.ParagraphFormat.SpaceAfter = 3
        sel.ParagraphFormat.KeepWithNext = False
        sel.ParagraphFormat.Alignment = 0

        # Render Question Number Prefix on the SAME line if present
        if q_num is not None:
            self.apply_native_numbered_list(word, sel, q_num=q_num, number_format=num_fmt)
            sel.Font.Bold = 0
        else:
            try:
                sel.Range.ListFormat.RemoveNumbers()
            except Exception:
                pass

        # Setup Tab Stops:
        # When Question prefix (e.g. 'Question 1.') is on the same line as options, offset tabs to balance all 4 columns cleanly
        if q_num is not None and cols > 1:
            bullet_sample = (num_fmt or "%1.").replace("%1", str(q_num or "1"))
            q_bullet_w_cm = max(0.8, (len(bullet_sample) * 0.19) + 0.35)
            rem_w = max(4.0, printable_width_cm - q_bullet_w_cm)
            col_slot = rem_w / cols
            tab_stops_cm = [q_bullet_w_cm + (col_slot * (i + 1)) for i in range(cols - 1)]
            sel.ParagraphFormat.TabStops.ClearAll()
            for t in tab_stops_cm:
                sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(t), Alignment=0)
        else:
            self.setup_tab_stops(sel, cols, left_indent_cm=left_indent_cm, printable_width_cm=printable_width_cm, max_item_len=max_item_len)


        opt_color_int = parse_color_to_rgb_int(self.opt_color)

        from uln_parser import parse_inline_spans
        for idx_item, (let_str, body_str) in enumerate(formatted_items):
            # Render Option Letter in BOLD with custom opt_color
            sel.Font.Name = self.font_name
            sel.Font.Size = self.font_size
            sel.Font.Bold = 1
            sel.Font.Italic = 0
            sel.Font.Underline = 0
            if opt_color_int is not None:
                sel.Font.Color = opt_color_int
            else:
                sel.Font.Color = 0
            sel.TypeText(f"{let_str} ")
            sel.Font.Bold = 0
            sel.Font.Color = 0  # Reset to black for body text

            # Render Option Body Text
            body_spans = parse_inline_spans(body_str)
            self.write_inline_spans(sel, body_spans, default_bold=False)

            # Insert tab or paragraph break based on calculated column layout
            if (idx_item + 1) % cols == 0 or idx_item == N - 1:
                sel.TypeParagraph()
            else:
                sel.TypeText("\t")

        sel.ParagraphFormat.LeftIndent = 0
        sel.ParagraphFormat.TabStops.ClearAll()
        self.last_rendered_tag = "OPT"


    def render_box_shape(self, sel, doc, word, block: ULNBlock, printable_width_cm: float):
        """
        Renders Word Bank / Callout Box with words typed directly INSIDE a MS Word Rounded Rectangle Shape TextFrame (msoShapeRoundedRectangle = 5).
        Anchors the box to paragraph line flow with wdWrapTopBottom so adding/deleting lines above moves the box AND its text together seamlessly.
        Configures symmetric margins and equal column slot widths for clean, balanced centering without excess space on the right.
        """
        printable_width_pt = cm_to_pt(printable_width_cm)
        
        if '|' in block.content:
            words = [w.strip() for w in block.content.split('|') if w.strip()]
        else:
            words = [w.strip() for w in block.content.split() if w.strip()]

        if not words:
            return

        N = len(words)

        char_w_pt = 6.2  # 12pt bold Times New Roman character width estimate
        margin_pt = 0.0  # 0mm inner margins on all 4 sides

        # Determine column count (1 to 5 cols)
        max_len_all = max(len(w) for w in words)
        est_slot_w = max(45.0, (max_len_all * char_w_pt) + 16.0)

        if est_slot_w >= (printable_width_pt - (2 * margin_pt)):
            cols = 1
        else:
            max_fit_cols = max(1, int((printable_width_pt - (2 * margin_pt)) / est_slot_w))
            if N <= 5:
                cols = min(N, max_fit_cols)
            elif N <= 8:
                cols = min(4, max_fit_cols)
            elif N <= 10:
                cols = min(5, max_fit_cols)
            else:
                cols = min(4, max_fit_cols)

        slot_w = max(45.0, (max_len_all * char_w_pt) + 16.0)
        inner_w = slot_w * cols
        margin_left_pt = cm_to_pt(0.2)  # 2mm = ~5.67 pt
        box_width_pt = min(printable_width_pt, inner_w + margin_left_pt)
        left_offset_pt = max(0.0, (printable_width_pt - box_width_pt) / 2.0)

        num_rows = math.ceil(N / cols)
        font_line_h = 16.0
        box_height_pt = (num_rows * font_line_h) + (2 * margin_pt) + 4.0

        p_anchor = doc.Range(sel.Range.Start, sel.Range.Start)
        try:
            p_anchor.ParagraphFormat.SpaceBefore = 14.0
            p_anchor.ParagraphFormat.SpaceAfter = 14.0
        except Exception:
            pass

        try:
            shape = doc.Shapes.AddShape(
                5,  # msoShapeRoundedRectangle = 5
                0,
                0,
                box_width_pt,
                box_height_pt,
                Anchor=p_anchor
            )
            shape.RelativeHorizontalPosition = 0  # wdRelativeHorizontalPositionMargin = 0
            shape.RelativeVerticalPosition = 2    # wdRelativeVerticalPositionParagraph = 2
            shape.Left = left_offset_pt
            shape.Top = 0
            shape.WrapFormat.Type = 7             # wdWrapInline = 7 ("In Line with Text")
            shape.WrapFormat.DistanceTop = 12.0
            shape.WrapFormat.DistanceBottom = 12.0

            tf = shape.TextFrame
            tf.MarginTop = 0.0
            tf.MarginBottom = 0.0
            tf.MarginLeft = margin_left_pt   # 2mm left margin
            tf.MarginRight = 0.0


            try:
                tf.AutoSize = False
            except Exception:
                pass

            shape.Fill.Visible = False  # Transparent fill
            shape.Line.Weight = 1.0     # 1pt rounded border
            shape.Line.ForeColor.RGB = 0  # Black border line

            # Select inside shape TextFrame to typeset text runs
            tf.TextRange.Select()
            box_sel = word.Selection
            box_sel.Font.Name = self.font_name
            box_sel.Font.Size = self.font_size
            box_sel.Font.Bold = 1
            box_sel.Font.Color = 0  # Pure Black RGB(0,0,0)

            box_sel.ParagraphFormat.SpaceBefore = 0
            box_sel.ParagraphFormat.SpaceAfter = 0
            box_sel.ParagraphFormat.LineSpacingRule = 0
            box_sel.ParagraphFormat.Alignment = 0  # Left align
            box_sel.ParagraphFormat.TabStops.ClearAll()

            for c in range(1, cols):
                box_sel.ParagraphFormat.TabStops.Add(Position=slot_w * c, Alignment=0)

            lines = []
            for i in range(0, N, cols):
                lines.append(words[i:i + cols])

            from uln_parser import parse_inline_spans
            for idx_line, chunk in enumerate(lines):
                if idx_line > 0:
                    box_sel.ParagraphFormat.SpaceBefore = 1.5

                box_sel.ParagraphFormat.SpaceAfter = 0

                for idx_w, word_txt in enumerate(chunk):
                    w_spans = parse_inline_spans(word_txt, default_bold=True)
                    self.write_inline_spans(box_sel, w_spans)
                    box_sel.Font.Color = 0  # Enforce black text
                    if idx_w < len(chunk) - 1:
                        box_sel.TypeText("\t")

                if idx_line < len(lines) - 1:
                    box_sel.TypeParagraph()

            # Convert shape to native InlineShape ("In Line with Text")
            try:
                shape.ConvertToInlineShape()
            except Exception:
                pass

        except Exception as e:
            print(f"[ULNRenderer] Warning creating TextFrame box shape: {e}")

        # Move selection back to document main story below the inline shape
        try:
            end_range = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
            end_range.Select()
            sel.ParagraphFormat.LeftIndent = 0
            sel.ParagraphFormat.RightIndent = 0
            sel.ParagraphFormat.Alignment = 1  # Center align the box on the page
            sel.ParagraphFormat.SpaceBefore = 14.0
            sel.ParagraphFormat.SpaceAfter = 14.0
            sel.TypeParagraph()
            sel.ParagraphFormat.SpaceBefore = 0
            sel.ParagraphFormat.SpaceAfter = 4
            sel.ParagraphFormat.Alignment = 0  # Reset to Left align for next text
        except Exception:
            pass

        self.last_rendered_tag = "BOX"






    def render_table(self, sel, doc, tdata, printable_width_cm: float, idx_block: int = 0, blocks: List[ULNBlock] = None):
        """
        Renders [TABLE] block structure:
        - If borderless: uses divided paragraph tab stop columns with custom spacing.
        - If bordered: uses a native MS Word table with cell margins & 1.0pt gridlines.
        - Empty/blank cells (______): ignores literal underscore text inside cell bodies so bottom border acts as answer line!
        """
        if not tdata.rows:
            return

        if self.last_rendered_tag == "BOX":
            sel.ParagraphFormat.SpaceBefore = 14
            sel.ParagraphFormat.SpaceAfter = 4
        else:
            sel.ParagraphFormat.SpaceBefore = 12
            sel.ParagraphFormat.SpaceAfter = 4

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
                txt_line = re.sub(r'#(\d+)', r'\1', txt_line)

                # Auto-prefix and format option letters for rows after row 0 in borderless table (e.g. A., B., C., D.)
                if idx_r >= 1 and len(text_rows) >= 3:
                    m_opt = re.match(r'^\s*(?:(?:\*\*|\*|\[|\(?)*([a-zA-Z][\.\)])(?:\*\*|\*|\]|\}|\{u\}|\))*)\s*(.*)$', txt_line)
                    if m_opt:
                        let_str = m_opt.group(1).upper().rstrip('.')
                        body_str = m_opt.group(2).strip()
                        txt_line = f"**{let_str}.** {body_str}"
                    else:
                        let_str = chr(65 + idx_r - 1)
                        txt_line = f"**{let_str}.** {txt_line.strip()}"

                is_opt_line = bool(re.match(r'^\s*\*?\*?[A-Da-d][\.\)]', txt_line)) or (idx_r >= 1 and len(text_rows) >= 3)
                left_ind_cm = 0.5 if is_opt_line else 0.0

                q_match = re.match(r'^\s*(?:#?(\d+)[\.\)]|Question\s+#?(\d+)[\.\)]?|Câu\s+#?(\d+)[\.\)]?)\s*(.*)$', txt_line, re.IGNORECASE) if idx_r == 0 else None

                if q_match and q_match.group(4).strip():
                    try:
                        sel.Range.ListFormat.ApplyNumberDefault()
                        sel.ParagraphFormat.LeftIndent = 0
                        sel.ParagraphFormat.FirstLineIndent = 0
                    except Exception:
                        pass
                    txt_line = q_match.group(4).strip()
                else:
                    try:
                        sel.Range.ListFormat.RemoveNumbers()
                    except Exception:
                        pass
                    sel.ParagraphFormat.LeftIndent = cm_to_pt(left_ind_cm)
                    sel.ParagraphFormat.FirstLineIndent = 0

                from uln_parser import parse_inline_spans
                spans = parse_inline_spans(txt_line)

                sel.ParagraphFormat.SpaceBefore = 14 if (idx_r == 0 and self.last_rendered_tag == "BOX") else 3
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
                sel.ParagraphFormat.SpaceBefore = 14 if (self.last_rendered_tag == "BOX") else 3
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

            # Set default 2mm cell margins (padding) on all 4 sides (top, bottom, left, right)
            margin_2mm_pt = cm_to_pt(0.2)  # 2mm = 0.2cm = ~5.67pt
            try:
                table.TopPadding = margin_2mm_pt
                table.BottomPadding = margin_2mm_pt
                table.LeftPadding = margin_2mm_pt
                table.RightPadding = margin_2mm_pt
            except Exception:
                pass

            # Identify columns with headers containing STT, NO, or NO. for center alignment
            center_col_indices = set()
            if tdata.rows:
                header_row = tdata.rows[0]
                for c_idx, h_cell in enumerate(header_row.cells):
                    clean_h = re.sub(r'[\*\_\`\[\]]', '', h_cell.content).strip().lower()
                    if clean_h in ["stt", "no", "no."]:
                        center_col_indices.add(c_idx + 1)

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
                        
                        is_center_col = (c_idx + 1) in center_col_indices
                        p_range.ParagraphFormat.Alignment = 1 if (cell.is_header or is_center_col) else 0

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
            sel.ParagraphFormat.SpaceBefore = 14 if (i == 0 and self.last_rendered_tag == "BOX") else 4
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
