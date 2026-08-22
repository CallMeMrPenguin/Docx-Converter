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

from renderer_utils import (
    cm_to_pt,
    pt_to_cm,
    COLOR_NAME_TO_RGB,
    HIGHLIGHT_NAME_TO_INDEX,
    parse_color_to_rgb_int,
    extract_question_prefix_and_body,
    split_line_into_option_items,
    apply_title_case_to_text,
    apply_sentence_case_to_text,
    SUPPORTED_IMAGE_EXTENSIONS,
    ensure_word_compatible_image
)
from renderer_blocks import RendererBlocksMixin


class ULNWordRenderer(RendererBlocksMixin):
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
        self.opt_color = self.settings.get("opt_color", self.settings.get("option_color", "#000000"))
        self.option_color = self.opt_color
        self.instruction_color = self.settings.get("instruction_color", self.settings.get("ins_color", "#000000"))

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
        """Returns next user-queued image in order, or falls back to test pic directory, converted if needed."""
        found_path = None
        if self.user_img_idx < len(self.user_images):
            imgPath = self.user_images[self.user_img_idx]
            self.user_img_idx += 1
            if os.path.exists(imgPath):
                found_path = imgPath

        if not found_path and pic and pic.filepath and os.path.exists(pic.filepath):
            found_path = pic.filepath

        if not found_path:
            test_pic_dir = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "test pic"))
            if os.path.exists(test_pic_dir):
                pics = [
                    os.path.join(test_pic_dir, f)
                    for f in os.listdir(test_pic_dir)
                    if f.lower().endswith(SUPPORTED_IMAGE_EXTENSIONS)
                ]
                if pics:
                    idx = abs(hash(pic.description if pic else "img")) % len(pics)
                    found_path = pics[idx]

        if found_path:
            return ensure_word_compatible_image(found_path)

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

    def write_inline_spans(self, sel, spans: List[InlineSpan], default_bold: bool = False, default_italic: bool = False, default_uppercase: bool = False, custom_font_size: Optional[float] = None, force_color: Optional[int] = None):
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
            sel.Font.Size = custom_font_size if custom_font_size is not None else self.font_size
            
            is_bold = span.bold or default_bold
            is_italic = span.italic or default_italic
            is_upper = span.uppercase or default_uppercase

            sel.Font.Bold = 1 if is_bold else 0
            sel.Font.Italic = 1 if is_italic else 0
            sel.Font.Underline = 1 if span.underline else 0

            if force_color is not None:
                try:
                    sel.Font.Color = force_color
                except Exception:
                    pass
            elif span.color:
                rgb_int = parse_color_to_rgb_int(span.color)
                if rgb_int is not None:
                    try:
                        sel.Font.Color = rgb_int
                    except Exception:
                        pass
            elif span.is_instruction and self.instruction_color:
                ins_color_int = parse_color_to_rgb_int(self.instruction_color)
                if ins_color_int is not None:
                    try:
                        sel.Font.Color = ins_color_int
                    except Exception:
                        pass
                else:
                    try:
                        sel.Font.Color = 0
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

            # If text contains <blank> or [BLANK], handle answer blank in pure black
            if re.search(r'<(?:blank|BLANK)>|\[(?:blank|BLANK)\]', text):
                parts = re.split(r'(<(?:blank|BLANK)>|\[(?:blank|BLANK)\])', text)
                for part in parts:
                    if re.match(r'^(?:<(?:blank|BLANK)>|\[(?:blank|BLANK)\])$', part, re.IGNORECASE):
                        sel.Font.Color = 0  # Enforce black for answer blank
                        sel.Font.Underline = 0
                        sel.TypeText('___________')
                    else:
                        if part:
                            p_txt = part.upper() if is_upper else part
                            sel.TypeText(p_txt)
            else:
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
            try:
                sel.Font.Color = 0
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
                sel.ParagraphFormat.PageBreakBefore = True
                sel.ParagraphFormat.Alignment = 1  # Centered
                self.write_inline_spans(sel, block.spans, default_bold=True, default_uppercase=True, custom_font_size=self.font_size + 1.0, force_color=0)
                sel.TypeParagraph()
                sel.ParagraphFormat.PageBreakBefore = False
                sel.ParagraphFormat.Alignment = 0

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
                sel.ParagraphFormat.PageBreakBefore = False
                sel.ParagraphFormat.Alignment = 0  # Left
                self.write_inline_spans(sel, block.spans, default_bold=True, default_uppercase=True, force_color=0)
                sel.TypeParagraph()
                sel.ParagraphFormat.KeepWithNext = False
                sel.ParagraphFormat.PageBreakBefore = False

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
                sel.ParagraphFormat.PageBreakBefore = False
                sel.ParagraphFormat.Alignment = 0  # Left
                for s in block.spans:
                    s.text = apply_title_case_to_text(s.text)
                self.write_inline_spans(sel, block.spans, default_bold=True, force_color=0)
                sel.TypeParagraph()
                sel.ParagraphFormat.KeepWithNext = False
                sel.ParagraphFormat.PageBreakBefore = False

            elif tag == "H4":
                try:
                    sel.Style = doc.Styles("Heading 4")
                except Exception:
                    pass
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.FirstLineIndent = 0
                sel.ParagraphFormat.SpaceBefore = 8
                sel.ParagraphFormat.SpaceAfter = 4
                sel.ParagraphFormat.KeepWithNext = True
                sel.ParagraphFormat.PageBreakBefore = False
                sel.ParagraphFormat.Alignment = 0  # Left
                for s in block.spans:
                    s.text = apply_sentence_case_to_text(s.text)
                self.write_inline_spans(sel, block.spans, default_bold=True, force_color=0)
                sel.TypeParagraph()
                sel.ParagraphFormat.KeepWithNext = False
                sel.ParagraphFormat.PageBreakBefore = False

            elif tag == "H5":
                try:
                    sel.Style = doc.Styles("Heading 5")
                except Exception:
                    pass
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.FirstLineIndent = 0
                sel.ParagraphFormat.SpaceBefore = 6
                sel.ParagraphFormat.SpaceAfter = 4
                sel.ParagraphFormat.KeepWithNext = True
                sel.ParagraphFormat.PageBreakBefore = False
                sel.ParagraphFormat.Alignment = 0  # Left
                for s in block.spans:
                    s.text = apply_sentence_case_to_text(s.text)
                self.write_inline_spans(sel, block.spans, default_bold=False, force_color=0)
                sel.TypeParagraph()
                sel.ParagraphFormat.KeepWithNext = False
                sel.ParagraphFormat.PageBreakBefore = False

            elif tag == "H6":
                try:
                    sel.Style = doc.Styles("Heading 6")
                except Exception:
                    pass
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.FirstLineIndent = 0
                sel.ParagraphFormat.SpaceBefore = 6
                sel.ParagraphFormat.SpaceAfter = 4
                sel.ParagraphFormat.KeepWithNext = True
                sel.ParagraphFormat.PageBreakBefore = False
                sel.ParagraphFormat.Alignment = 0  # Left
                for s in block.spans:
                    s.text = apply_sentence_case_to_text(s.text)
                self.write_inline_spans(sel, block.spans, default_bold=False, default_italic=True, force_color=0)
                sel.TypeParagraph()
                sel.ParagraphFormat.KeepWithNext = False
                sel.ParagraphFormat.PageBreakBefore = False

            elif tag in ["P0", "P"]:
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.FirstLineIndent = 0
                sel.ParagraphFormat.PageBreakBefore = False
                is_ins_block = block.is_instruction or any(s.is_instruction for s in block.spans)
                if is_ins_block:
                    sel.ParagraphFormat.SpaceBefore = 14 if self.last_rendered_tag == "BOX" else 8
                    sel.ParagraphFormat.SpaceAfter = 4
                    sel.ParagraphFormat.KeepWithNext = True
                else:
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
                    sel.Font.Color = 0  # Enforce black for answer blank leader
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
                    sel.Font.Color = 0  # Enforce black for trailing answer blank
                    sel.Font.Underline = 0
                    sel.Font.Bold = 0
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
                    sel.Font.Color = 0  # Enforce black for answer blank leader
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
                    sel.Font.Color = 0  # Enforce black for trailing answer blank
                    sel.Font.Underline = 0
                    sel.Font.Bold = 0
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
                sel.ParagraphFormat.KeepWithNext = False
                sel.ParagraphFormat.PageBreakBefore = False

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
                    # 1. Optimal col2 start position based on longest Column 1 sentence (avg 0.155cm per char for 12pt Times New Roman)
                    est_c1_w = base_indent_cm + (max_c1_len * 0.155) + 0.4
                    col2_tab_pos_cm = max(base_indent_cm + 8.0, est_c1_w)
                    col2_tab_pos_cm = min(col2_tab_pos_cm, printable_width_cm - 2.5)

                    # 2. Safe blank width: strictly bounded within printable width with 0.4cm buffer to guarantee 0 line wraps
                    avail_w_cm = max(1.5, printable_width_cm - col2_tab_pos_cm - 0.4)
                    blank_w_cm = min(3.2, avail_w_cm)
                    char_under_w_cm = max(0.18, (self.font_size * 0.44) / 28.3465)
                    num_underscores = max(6, int(blank_w_cm / char_under_w_cm))

                    # Configure paragraph & tab stop before writing text
                    sel.ParagraphFormat.LeftIndent = cm_to_pt(base_indent_cm)
                    sel.ParagraphFormat.FirstLineIndent = 0
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(col2_tab_pos_cm), Alignment=0)
                    sel.ParagraphFormat.KeepWithNext = False
                    sel.ParagraphFormat.PageBreakBefore = False

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

                    sel.Font.Color = 0
                    sel.Font.Bold = 0
                    sel.Font.Underline = 0
                    sel.TypeText(f"\t{'_' * num_underscores}")
                    sel.TypeParagraph()

                else:
                    col2_trim = block.col2.strip()
                    m_opt = re.match(r'^\s*(?:(?:\*\*|\*|\[|\(?)*([a-zA-Z])[\.\)](?:\*\*|\*|\]|\}|\{u\}|\))*)\s+(.*)$', col2_trim)

                    pref, delim, q_num, c1_body = extract_question_prefix_and_body(block.col1)
                    if q_num is not None:
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

                    # Single tab stop: col1 text → \t → col2 content (option letter + body inline)
                    sel.ParagraphFormat.LeftIndent = cm_to_pt(col2_tab_pos_cm)
                    sel.ParagraphFormat.FirstLineIndent = cm_to_pt(-col1_needed_cm)
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(col2_tab_pos_cm), Alignment=0)
                    sel.ParagraphFormat.KeepWithNext = False
                    sel.ParagraphFormat.PageBreakBefore = False

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
