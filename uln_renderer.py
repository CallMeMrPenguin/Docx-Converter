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

try:
    import ctypes
    _user32 = ctypes.windll.user32
    _get_async_key_state = _user32.GetAsyncKeyState
except Exception:
    _get_async_key_state = None

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
    ensure_word_compatible_image,
    natural_sort_key
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
        - For parenthesis prefixes (e.g. '(' -> '(%1)'), preserves tight bracket fit without unwanted space.
        - If delimiter was selected in GUI settings (and is non-default, e.g. ':', ')', '-'), use GUI delimiter.
          Otherwise use extracted delimiter from text or global default.
        """
        raw_pref = extracted_pref if (extracted_pref and extracted_pref.strip()) else (self.question_prefix or "")
        if raw_pref and raw_pref.strip():
            p_strip = raw_pref.strip()
            if p_strip.endswith("("):
                pref = p_strip
            else:
                pref = f"{p_strip} "
        else:
            pref = ""

        if self.question_delimiter and self.question_delimiter != "." and "(" not in pref:
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
                    pics.sort(key=lambda p: natural_sort_key(os.path.basename(p)))
                    idx = abs(hash(pic.description if pic else "img")) % len(pics)
                    found_path = pics[idx]

        if found_path:
            return ensure_word_compatible_image(found_path)

        return None

    def check_cancellation(self):
        """Checks if user pressed the ESC key anywhere on the system to instantly cancel generation."""
        if _get_async_key_state and (_get_async_key_state(0x1B) & 0x8000):
            raise KeyboardInterrupt("Tác vụ tạo DOCX đã bị người dùng hủy bằng phím ESC.")

    def configure_document(self, doc):
        """Applies page setup margins."""
        self.check_cancellation()
        ps = doc.PageSetup
        ps.PageWidth = cm_to_pt(21.0)
        ps.PageHeight = cm_to_pt(29.7)
        ps.TopMargin = cm_to_pt(self.margin_top)
        ps.BottomMargin = cm_to_pt(self.margin_bottom)
        ps.LeftMargin = cm_to_pt(self.margin_left)
        ps.RightMargin = cm_to_pt(self.margin_right)

    def apply_page_numbers(self, doc):
        """Applies footer page numbering AFTER all content has finished rendering to avoid rendering lag."""
        if not self.enable_page_numbers:
            return

        try:
            self.check_cancellation()
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

            try:
                doc.Fields.Update()
            except Exception:
                pass
        except Exception as e:
            print(f"[ULNRenderer] Warning applying page numbers: {e}")

    def write_inline_spans(self, sel, spans: List[InlineSpan], default_bold: bool = False, default_italic: bool = False, default_uppercase: bool = False, custom_font_size: Optional[float] = None, force_color: Optional[int] = None):
        """Writes formatted text runs strictly according to span AST properties with cached COM attributes for 10x speed."""
        f = sel.Font
        f_size = custom_font_size if custom_font_size is not None else self.font_size
        
        for idx, span in enumerate(spans):
            text = span.text

            # Check if span is an inline [PIC...] tag
            if text.startswith("[PIC:") or text.strip().upper() == "[PIC]":
                pic_info = parse_pic_tag(text) or PicInfo(description="Activity Picture", pos="center", size="small")
                self.render_pic(sel, None, pic_info)
                # Only add trailing space if followed by more non-whitespace text in this span list
                if idx + 1 < len(spans) and spans[idx + 1].text.strip():
                    sel.TypeText(" ")
                continue

            is_bold = 1 if (span.bold or default_bold) else 0
            is_italic = 1 if (span.italic or default_italic) else 0
            is_upper = span.uppercase or default_uppercase
            is_under = 1 if span.underline else 0

            # Direct minimal COM property assignment
            f.Name = self.font_name
            f.Size = f_size
            f.Bold = is_bold
            f.Italic = is_italic
            f.Underline = is_under

            if force_color is not None:
                f.Color = force_color
            elif span.color:
                rgb_int = parse_color_to_rgb_int(span.color)
                f.Color = rgb_int if rgb_int is not None else 0
            elif span.is_instruction and self.instruction_color:
                ins_color_int = parse_color_to_rgb_int(self.instruction_color)
            else:
                f.Color = 0

            if span.bg_color:
                try:
                    sel.Range.HighlightColorIndex = HIGHLIGHT_NAME_TO_INDEX.get(span.bg_color.lower(), 7)
                except Exception:
                    pass

            # Standardize excessive underscores (>45) to an optimal blank line length
            if re.match(r'^_{30,}$', text):
                text = '_' * 35

            # If text contains (number), (#number), [number], [#number], #number, or <blank>/[BLANK], format them cleanly inline
            inline_token_pat = r'(\(\s*#?\d+\s*[\.\:\)]*\)|\[\s*#?\d+\s*[\.\:\)]*\]|(?:^|(?<=\s))#\s*\d+[\.\:\)\/\-]*(?=\s|$)|<(?:blank|BLANK)>|\[(?:blank|BLANK)\])'
            if re.search(inline_token_pat, text):
                parts = re.split(inline_token_pat, text)
                base_color = force_color if force_color is not None else (parse_color_to_rgb_int(span.color) if span.color else (parse_color_to_rgb_int(self.instruction_color) if (span.is_instruction and self.instruction_color) else 0))
                for part in parts:
                    if not part:
                        continue
                    if re.match(r'^\(\s*#?\d+\s*[\.\:\)]*\)$', part):
                        clean_paren = re.sub(r'#|\s', '', part)
                        f.Bold = 1
                        q_col = parse_color_to_rgb_int(self.question_color)
                        f.Color = q_col if q_col is not None else 0
                        sel.TypeText(clean_paren)
                        f.Bold = is_bold
                        f.Color = base_color if base_color is not None else 0
                    elif re.match(r'^\[\s*#?\d+\s*[\.\:\)]*\]$', part):
                        clean_bracket = re.sub(r'#|\s', '', part)
                        f.Bold = 1
                        q_col = parse_color_to_rgb_int(self.question_color)
                        f.Color = q_col if q_col is not None else 0
                        sel.TypeText(clean_bracket)
                        f.Bold = is_bold
                        f.Color = base_color if base_color is not None else 0
                    elif re.match(r'^#\s*\d+[\.\:\)\/\-]*$', part):
                        clean_num = re.sub(r'^#\s*', '', part)
                        f.Bold = 1
                        q_col = parse_color_to_rgb_int(self.question_color)
                        f.Color = q_col if q_col is not None else 0
                        sel.TypeText(clean_num)
                        f.Bold = is_bold
                        f.Color = base_color if base_color is not None else 0
                    elif re.match(r'^(?:<(?:blank|BLANK)>|\[(?:blank|BLANK)\])$', part, re.IGNORECASE):
                        f.Color = 0
                        f.Underline = 0
                        sel.TypeText('___________')
                    else:
                        sel.TypeText(part.upper() if is_upper else part)
            else:
                sel.TypeText(text.upper() if is_upper else text)

        try:
            f.Underline = 0
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

    def is_in_multiple_choice_question(self, blocks: List[ULNBlock], curr_idx: int) -> bool:
        """
        Scans forward from curr_idx to determine if an [OPT] block follows within the current question/dialogue.
        Stops scanning when a new numbered question, section header (H1-H6), BOX, TABLE, or PIC_GRID is reached.
        """
        for k in range(curr_idx + 1, len(blocks)):
            b = blocks[k]
            if b.tag == "OPT":
                return True
            if b.tag in ("H1", "H2", "H3", "H4", "H5", "H6", "BOX", "TABLE", "PIC_GRID"):
                break
            if b.tag == "P0":
                pref, delim, q_num, _ = extract_question_prefix_and_body(b.content)
                if q_num is not None:
                    break
        return False

    def render(self, blocks: List[ULNBlock], doc, word, is_root: bool = True):
        """Renders parsed ULNBlocks into the active document purely driven by structural AST tags."""
        if is_root:
            self.configure_document(doc)
        sel = word.Selection

        printable_width_cm = 21.0 - self.margin_left - self.margin_right

        idx_block = 0
        while idx_block < len(blocks):
            self.check_cancellation()
            sel = word.Selection
            block = blocks[idx_block]
            tag = block.tag

            sel.Font.Italic = 0
            sel.Font.Bold = 0
            sel.Font.Underline = 0
            sel.Font.Color = 0

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
                    self.current_exercise_q_num = 1
                    self._exercise_list_template = None
                    self.is_first_question_in_num_block = True
                
                # Check if this P0 is followed by an OPT block, Dialogue continuation, BOX, or Sentence Rewrite blank line
                has_next_opt = self.is_in_multiple_choice_question(blocks, idx_block)
                has_next_box = (idx_block + 1 < len(blocks) and blocks[idx_block + 1].tag == "BOX")
                has_next_dlg = (idx_block + 1 < len(blocks) and blocks[idx_block + 1].tag in ["P1", "P0"] and bool(re.search(r'^\s*(?:(?:\*\*|\*|\[)?(?:Speaker\s+)?[A-Za-z0-9]+\s*[:\.\-](?:\*\*|\*|\])?)\s*', blocks[idx_block + 1].content, re.IGNORECASE)))
                has_next_rewrite_blank = (idx_block + 1 < len(blocks) and blocks[idx_block + 1].tag in ["P1", "P2"] and (
                    bool(re.search(r'<(?:blank|BLANK)>|\[(?:blank|BLANK)\]|_{3,}|(?:→|->)', blocks[idx_block + 1].content))
                ))
                pref_chk, delim_chk, q_num_chk, body_chk = extract_question_prefix_and_body(block.content)
                is_numbered_q = (q_num_chk is not None)
                is_dialogue_line = bool(re.search(r'(?:^|#\d+[\.\)]\s*)(?:(?:\*\*|\*|\[)?(?:Speaker\s+)?[A-Za-z0-9]+\s*[:\.\-](?:\*\*|\*|\])?)\s*', block.content, re.IGNORECASE))

                if getattr(self, "is_inside_num_container", False) and tag != "PIC_GRID":
                    sel.ParagraphFormat.SpaceBefore = 2
                    sel.ParagraphFormat.SpaceAfter = 2
                    sel.ParagraphFormat.LineSpacing = self.font_size * 1.16
                    sel.ParagraphFormat.KeepWithNext = bool(is_numbered_q and (has_next_opt or has_next_dlg or has_next_rewrite_blank))
                elif is_ins_block or (is_numbered_q and (has_next_opt or has_next_dlg or has_next_rewrite_blank)) or has_next_box:
                    sel.ParagraphFormat.SpaceBefore = 14 if self.last_rendered_tag == "BOX" else (8 if is_ins_block else 6)
                    sel.ParagraphFormat.SpaceAfter = 4 if is_ins_block else 2
                    sel.ParagraphFormat.KeepWithNext = True
                else:
                    sel.ParagraphFormat.SpaceBefore = 6
                    sel.ParagraphFormat.SpaceAfter = 4
                    sel.ParagraphFormat.KeepWithNext = False
                sel.ParagraphFormat.Alignment = 0

                # Check if paragraph is ONLY a standalone blank line _____ or <blank> or [BLANK], optionally followed by symbol/punct
                blank_symbol_match = re.match(r'^\s*(?:_{3,}|<(?:blank|BLANK)>|\[(?:blank|BLANK)\])\s*([?\.\!:,;]?)\s*$', block.content, re.IGNORECASE)
                trailing_pic_match = re.search(r'\s*(\[PIC(?::[^\]]+)?\])\s*$', block.content, re.IGNORECASE)
                is_arrow_rewrite = bool(re.search(r'(?:→|->|=>|➔|➜)', block.content))
                # STRICT RULE: NEVER enable full tab leader if has_next_opt is True!
                # ONLY enable for explicit arrow rewrites (→), standalone long blank rewrites (____), or non-OPT dialogues!
                should_allow_full_blank = (not has_next_opt) and (is_arrow_rewrite or bool(re.search(r'_{15,}', block.content)) or (is_dialogue_line and not has_next_opt))
                trailing_blank_symbol_match = re.match(r'^(.+?)\s*(?:<(?:blank|BLANK)>|\[(?:blank|BLANK)\]|_{3,})\s*([?\.\!:,;]?)\s*$', block.content, re.DOTALL | re.IGNORECASE) if (not blank_symbol_match and not trailing_pic_match and should_allow_full_blank) else None

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
                elif trailing_pic_match:
                    text_part = block.content[:trailing_pic_match.start()].strip()
                    pic_str = trailing_pic_match.group(1).strip()
                    pref, delim, q_num, c_body = extract_question_prefix_and_body(text_part)

                    if has_next_opt:
                        tab_fake = ULNBlock(tag="TAB2", col1=text_part, col2=pic_str)
                        self.render_side_by_side_pic_mcq(sel, doc, word, tab_fake, blocks[idx_block + 1], printable_width_cm)
                        idx_block += 2
                        continue

                    pic_info = parse_pic_tag(pic_str) or PicInfo(description="Sign / Picture", pos="center", size="small")

                    if q_num is not None and not c_body.strip():
                        # Question is purely a number and picture: e.g. #1. [PIC]
                        sel.ParagraphFormat.LeftIndent = 0
                        sel.ParagraphFormat.FirstLineIndent = 0
                        sel.ParagraphFormat.TabStops.ClearAll()
                        sel.ParagraphFormat.SpaceBefore = 6 if self.last_rendered_tag == "OPT" else 3
                        sel.ParagraphFormat.SpaceAfter = 3
                        sel.ParagraphFormat.KeepWithNext = True

                        pref_str = pref if pref else ""
                        delim_char = delim if delim else "."
                        num_prefix_str = f"{pref_str}{q_num}{delim_char} "
                        sel.Font.Name = self.font_name
                        sel.Font.Size = self.font_size
                        sel.Font.Bold = 1
                        sel.Font.Italic = 0
                        sel.Font.Underline = 0
                        q_color_int = parse_color_to_rgb_int(self.question_color)
                        sel.Font.Color = q_color_int if q_color_int is not None else 0
                        sel.TypeText(num_prefix_str)
                        sel.Font.Bold = 0
                        sel.Font.Color = 0

                        pic_w_cm = 2.2
                        pic_h_cm = 2.2
                        self.current_tab2_pic_width_cm = pic_w_cm
                        self.current_tab2_pic_height_cm = pic_h_cm
                        self.render_pic(sel, doc, pic_info)
                        self.current_tab2_pic_width_cm = None
                        self.current_tab2_pic_height_cm = None
                        sel.TypeParagraph()
                    else:
                        pic_w_cm = 2.0
                        pic_h_cm = 1.4
                        col_pic_pos_cm = printable_width_cm - pic_w_cm

                        sel.ParagraphFormat.LeftIndent = 0
                        sel.ParagraphFormat.FirstLineIndent = 0
                        sel.ParagraphFormat.TabStops.ClearAll()
                        sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(col_pic_pos_cm), Alignment=0)
                        sel.ParagraphFormat.SpaceBefore = 3
                        sel.ParagraphFormat.SpaceAfter = 3
                        sel.ParagraphFormat.KeepWithNext = True

                        if q_num is not None and c_body.strip():
                            num_fmt = self.get_effective_number_format(pref, delim)
                            self.apply_native_numbered_list(word, sel, q_num=q_num, number_format=num_fmt)
                            from uln_parser import parse_inline_spans as _pis
                            text_spans = _pis(c_body.strip())
                            self.write_inline_spans(sel, text_spans)
                        else:
                            try:
                                sel.Range.ListFormat.RemoveNumbers()
                            except Exception:
                                pass
                            from uln_parser import parse_inline_spans as _pis
                            text_spans = _pis(text_part)
                            self.write_inline_spans(sel, text_spans)

                        sel.TypeText("\t")
                        self.current_tab2_pic_width_cm = pic_w_cm
                        self.current_tab2_pic_height_cm = pic_h_cm
                        self.render_pic(sel, doc, pic_info)
                        self.current_tab2_pic_width_cm = None
                        self.current_tab2_pic_height_cm = None
                        sel.TypeParagraph()
                        sel.ParagraphFormat.TabStops.ClearAll()
                elif trailing_blank_symbol_match:
                    # Option B: text before <blank> rendered inline, blank filled dynamically to right margin via Leader=4
                    text_part = trailing_blank_symbol_match.group(1)
                    trailing_sym = trailing_blank_symbol_match.group(2).strip()

                    pref, delim, q_num, c_body = extract_question_prefix_and_body(text_part)
                    if q_num is not None and c_body.strip():
                        if is_dialogue_line:
                            try:
                                sel.Range.ListFormat.RemoveNumbers()
                            except Exception:
                                pass
                            sel.ParagraphFormat.TabStops.ClearAll()
                            sel.ParagraphFormat.LeftIndent = 0
                            sel.ParagraphFormat.FirstLineIndent = 0
                            sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(0.63), Alignment=0)
                            sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(printable_width_cm), Alignment=2, Leader=4)
                            
                            sel.Font.Name = self.font_name
                            sel.Font.Size = self.font_size
                            sel.Font.Bold = 1
                            q_col = parse_color_to_rgb_int(self.question_color)
                            sel.Font.Color = q_col if q_col is not None else 0
                            num_fmt = self.get_effective_number_format(pref, delim)
                            num_str = num_fmt.replace("%1", str(q_num)) if "%1" in num_fmt else f"{q_num}."
                            sel.TypeText(num_str)
                            sel.TypeText("\t")
                            sel.Font.Bold = 0
                            sel.Font.Color = 0
                            text_part = c_body.strip()
                        else:
                            num_fmt = self.get_effective_number_format(pref, delim)
                            self.apply_native_numbered_list(word, sel, q_num=q_num, number_format=num_fmt)
                            if self.last_rendered_tag == "BOX":
                                sel.ParagraphFormat.SpaceBefore = 14
                            text_part = c_body.strip()
                            sel.ParagraphFormat.TabStops.ClearAll()
                            sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(printable_width_cm), Alignment=2, Leader=4)
                    else:
                        q_num_match = re.match(r'^\s*(?:#?(\d+)[\.\)]|Question\s+#?(\d+)[\.\)]?|Câu\s+#?(\d+)[\.\)]?)\s*(.*)$', text_part, re.IGNORECASE)
                        if q_num_match and q_num_match.group(4).strip():
                            q_num_val = q_num_match.group(1) or q_num_match.group(2) or q_num_match.group(3)
                            if is_dialogue_line:
                                try:
                                    sel.Range.ListFormat.RemoveNumbers()
                                except Exception:
                                    pass
                                sel.ParagraphFormat.TabStops.ClearAll()
                                sel.ParagraphFormat.LeftIndent = 0
                                sel.ParagraphFormat.FirstLineIndent = 0
                                sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(0.63), Alignment=0)
                                sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(printable_width_cm), Alignment=2, Leader=4)
                                sel.Font.Name = self.font_name
                                sel.Font.Size = self.font_size
                                sel.Font.Bold = 1
                                q_col = parse_color_to_rgb_int(self.question_color)
                                sel.Font.Color = q_col if q_col is not None else 0
                                sel.TypeText(f"{q_num_val}.")
                                sel.TypeText("\t")
                                sel.Font.Bold = 0
                                sel.Font.Color = 0
                                text_part = q_num_match.group(4).strip()
                            else:
                                self.apply_native_numbered_list(word, sel, q_num=q_num_val)
                                if self.last_rendered_tag == "BOX":
                                    sel.ParagraphFormat.SpaceBefore = 14
                                text_part = q_num_match.group(4).strip()
                                sel.ParagraphFormat.TabStops.ClearAll()
                                sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(printable_width_cm), Alignment=2, Leader=4)
                        else:
                            try:
                                sel.Range.ListFormat.RemoveNumbers()
                            except Exception:
                                pass
                            if self.last_rendered_tag == "BOX":
                                sel.ParagraphFormat.SpaceBefore = 14
                            sel.ParagraphFormat.TabStops.ClearAll()
                            sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(printable_width_cm), Alignment=2, Leader=4)

                    from uln_parser import parse_inline_spans as _pis
                    text_spans = _pis(text_part)
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
                        if is_dialogue_line:
                            try:
                                sel.Range.ListFormat.RemoveNumbers()
                            except Exception:
                                pass
                            sel.ParagraphFormat.TabStops.ClearAll()
                            sel.ParagraphFormat.LeftIndent = 0
                            sel.ParagraphFormat.FirstLineIndent = 0
                            sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(0.63), Alignment=0)
                            sel.Font.Name = self.font_name
                            sel.Font.Size = self.font_size
                            sel.Font.Bold = 1
                            q_col = parse_color_to_rgb_int(self.question_color)
                            sel.Font.Color = q_col if q_col is not None else 0
                            num_fmt = self.get_effective_number_format(pref, delim)
                            num_str = num_fmt.replace("%1", str(q_num)) if "%1" in num_fmt else f"{q_num}."
                            sel.TypeText(num_str)
                            sel.TypeText("\t")
                            sel.Font.Bold = 0
                            sel.Font.Color = 0
                        else:
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
                        opt_color_int = parse_color_to_rgb_int(self.opt_color)
                        for idx_item, item in enumerate(items):
                            m_let = re.match(r'^\s*(?:(?:\*\*|\*|\[|\(?)*([a-zA-Z0-9][\.\)])(?:\*\*|\*|\]|\}|\{u\}|\))*)\s+(.*)$', item.strip())
                            if m_let:
                                let_part = m_let.group(1).rstrip('.)')
                                body_part = m_let.group(2).strip()
                                sel.Font.Name = self.font_name
                                sel.Font.Size = self.font_size
                                sel.Font.Bold = 1
                                sel.Font.Italic = 0
                                sel.Font.Underline = 0
                                sel.Font.Color = opt_color_int if opt_color_int is not None else 0
                                sel.TypeText(f"{let_part}. ")
                                sel.Font.Bold = 0
                                sel.Font.Color = 0
                                spans = parse_inline_spans(body_part)
                                self.write_inline_spans(sel, spans)
                            else:
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
                is_dlg_speaker = bool(re.match(r'^\s*(?:(?:\*\*|\*|\[)?(?:Speaker\s+)?[A-Za-z0-9]+\s*[:\.\-](?:\*\*|\*|\])?)\s*', content_to_render, re.IGNORECASE))

                if q_num is not None:
                    left_indent_cm = 0.0
                else:
                    left_indent_cm = 0.5 if tag == "P1" else 1.0

                if q_num is not None:
                    num_fmt = self.get_effective_number_format(pref, delim)
                    self.apply_native_numbered_list(word, sel, q_num=q_num, number_format=num_fmt)
                else:
                    try:
                        sel.Range.ListFormat.RemoveNumbers()
                    except Exception:
                        pass

                has_next_dlg = (idx_block + 1 < len(blocks) and blocks[idx_block + 1].tag in ["P1", "P0"] and bool(re.match(r'^\s*(?:(?:\*\*|\*|\[)?(?:Speaker\s+)?[A-Za-z0-9]+\s*[:\.\-](?:\*\*|\*|\])?)\s*', blocks[idx_block + 1].content, re.IGNORECASE)))
                has_next_opt = self.is_in_multiple_choice_question(blocks, idx_block)

                if getattr(self, "is_inside_num_container", False) and tag != "PIC_GRID":
                    sel.ParagraphFormat.SpaceBefore = 2
                    sel.ParagraphFormat.SpaceAfter = 2
                    sel.ParagraphFormat.LineSpacing = self.font_size * 1.16
                    sel.ParagraphFormat.KeepWithNext = False
                elif is_dlg_speaker:
                    sel.ParagraphFormat.SpaceBefore = 1
                    sel.ParagraphFormat.SpaceAfter = 2 if has_next_dlg else 6
                    sel.ParagraphFormat.KeepWithNext = has_next_dlg
                else:
                    space_before_p1 = 14 if (self.last_rendered_tag == "BOX") else (4 if tag == "P1" else 3)
                    sel.ParagraphFormat.SpaceBefore = space_before_p1
                    sel.ParagraphFormat.SpaceAfter = 3
                    sel.ParagraphFormat.KeepWithNext = False

                sel.ParagraphFormat.Alignment = 0

                items = split_line_into_option_items(content_to_render)

                # Check if paragraph is ONLY a standalone blank line _____ or <blank> or [BLANK], optionally followed by symbol/punct
                blank_symbol_match = re.match(r'^\s*(?:_{3,}|<(?:blank|BLANK)>|\[(?:blank|BLANK)\])\s*([?\.\!:,;]?)\s*$', content_to_render, re.IGNORECASE)
                # Check if paragraph has trailing picture [PIC] or [PIC: ...]
                trailing_pic_match = re.search(r'\s*(\[PIC(?::[^\]]+)?\])\s*$', content_to_render, re.IGNORECASE)
                is_arrow_rewrite = bool(re.search(r'(?:→|->|=>|➔|➜)', content_to_render)) or bool(re.search(r'(?:→|->|=>|➔|➜)', block.content))
                is_arrow_only = bool(re.match(r'^\s*(?:→|->|=>|➔|➜)\s*$', content_to_render))
                should_allow_full_blank = (not has_next_opt) and (is_arrow_rewrite or is_arrow_only or (tag == "P1") or bool(re.search(r'_{15,}', content_to_render)) or is_dlg_speaker)
                trailing_blank_symbol_match = re.match(r'^(.+?)\s*(?:<(?:blank|BLANK)>|\[(?:blank|BLANK)\]|_{3,})\s*([?\.\!:,;]?)\s*$', content_to_render, re.DOTALL | re.IGNORECASE) if (not blank_symbol_match and not trailing_pic_match and should_allow_full_blank) else None

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
                elif trailing_pic_match:
                    text_part = content_to_render[:trailing_pic_match.start()].strip()
                    pic_str = trailing_pic_match.group(1).strip()

                    if has_next_opt:
                        tab_fake = ULNBlock(tag="TAB2", col1=text_part, col2=pic_str)
                        self.render_side_by_side_pic_mcq(sel, doc, word, tab_fake, blocks[idx_block + 1], printable_width_cm)
                        idx_block += 2
                        continue

                    pic_info = parse_pic_tag(pic_str) or PicInfo(description="Activity Picture", pos="right", size="small")
                    pic_w_cm = 2.0
                    pic_h_cm = 1.4
                    col_pic_pos_cm = printable_width_cm - pic_w_cm

                    sel.ParagraphFormat.LeftIndent = cm_to_pt(left_indent_cm)
                    sel.ParagraphFormat.FirstLineIndent = 0
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(col_pic_pos_cm), Alignment=0)

                    from uln_parser import parse_inline_spans as _pis
                    text_spans = _pis(text_part)
                    self.write_inline_spans(sel, text_spans)

                    sel.TypeText("\t")
                    self.current_tab2_pic_width_cm = pic_w_cm
                    self.current_tab2_pic_height_cm = pic_h_cm
                    self.render_pic(sel, doc, pic_info)
                    self.current_tab2_pic_width_cm = None
                    self.current_tab2_pic_height_cm = None
                    sel.TypeParagraph()
                    sel.ParagraphFormat.TabStops.ClearAll()
                elif trailing_blank_symbol_match or is_arrow_only:
                    text_part = trailing_blank_symbol_match.group(1) if trailing_blank_symbol_match else content_to_render.strip()
                    trailing_sym = trailing_blank_symbol_match.group(2).strip() if trailing_blank_symbol_match else ""

                    sel.ParagraphFormat.LeftIndent = cm_to_pt(left_indent_cm)
                    sel.ParagraphFormat.FirstLineIndent = 0
                    sel.ParagraphFormat.TabStops.ClearAll()
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(printable_width_cm), Alignment=2, Leader=4)

                    pref, delim, q_num, c_body = extract_question_prefix_and_body(text_part)
                    if q_num is not None and c_body.strip():
                        pref_str = pref if pref else ""
                        delim_char = delim if delim else "."
                        num_prefix_str = f"{pref_str}{q_num}{delim_char} "
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
                        sel.TypeText(num_prefix_str)
                        sel.Font.Bold = 0
                        sel.Font.Color = 0
                        text_part = c_body.strip()

                    from uln_parser import parse_inline_spans as _pis
                    text_spans = _pis(text_part)
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
                            opt_color_int = parse_color_to_rgb_int(self.opt_color)
                            for idx_item, item in enumerate(items):
                                m_let = re.match(r'^\s*(?:(?:\*\*|\*|\[|\(?)*([a-zA-Z0-9][\.\)])(?:\*\*|\*|\]|\}|\{u\}|\))*)\s+(.*)$', item.strip())
                                if m_let:
                                    let_part = m_let.group(1).rstrip('.)')
                                    body_part = m_let.group(2).strip()
                                    sel.Font.Name = self.font_name
                                    sel.Font.Size = self.font_size
                                    sel.Font.Bold = 1
                                    sel.Font.Italic = 0
                                    sel.Font.Underline = 0
                                    sel.Font.Color = opt_color_int if opt_color_int is not None else 0
                                    sel.TypeText(f"{let_part}. ")
                                    sel.Font.Bold = 0
                                    sel.Font.Color = 0
                                    spans = parse_inline_spans(body_part)
                                    self.write_inline_spans(sel, spans)
                                else:
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

            elif tag == "TABLE":
                if block.table_data:
                    self.render_table(sel, doc, block.table_data, printable_width_cm, idx_block=idx_block, blocks=blocks)

            elif tag.startswith("TAB") and tag != "TABLE":
                num_cols = len(block.cols) if block.cols else (3 if tag == "TAB3" else (4 if tag == "TAB4" else 2))
                if num_cols >= 3:
                    self.current_block_tag = tag
                    self.render_tab_multi(sel, doc, block, idx_block, blocks, printable_width_cm)
                else:
                    self.current_block_tag = "TAB2"
                    tab2_group = []
                    k = idx_block
                    while k < len(blocks) and blocks[k].tag.startswith("TAB") and blocks[k].tag != "TABLE" and (not blocks[k].cols or len(blocks[k].cols) == 2):
                        tab2_group.append(blocks[k])
                        k += 1

                    # Check if single TAB2 block contains a picture and is immediately followed by an OPT block
                    if len(tab2_group) == 1:
                        single_tab = tab2_group[0]
                        c1_txt = single_tab.col1 or ""
                        c2_txt = single_tab.col2 or ""
                        has_pic_in_tab = ("[PIC" in c1_txt.upper() or "[PIC" in c2_txt.upper() or parse_pic_tag(c1_txt) is not None or parse_pic_tag(c2_txt) is not None)
                        if has_pic_in_tab and k < len(blocks) and blocks[k].tag == "OPT":
                            self.render_side_by_side_pic_mcq(sel, doc, word, single_tab, blocks[k], printable_width_cm)
                            idx_block = k + 1
                            continue

                    self.render_tab2_group(sel, doc, word, tab2_group, printable_width_cm)
                    idx_block += len(tab2_group)
                    continue

            elif tag == "PIC_GRID":
                self.render_pic_grid(sel, doc, block.children, printable_width_cm)

            elif tag == "BOX":
                self.render_box_shape(sel, doc, word, block, printable_width_cm, idx_block=idx_block, blocks=blocks)

            elif tag == "NUM":
                self.render_num_container(sel, doc, word, block, printable_width_cm)

            elif tag == "OPT":
                if idx_block == 0 or blocks[idx_block - 1].tag != "OPT":
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
                raw_c = block.content.strip()
                sub = getattr(block, "sub_tag", None)

                # Determine if this line is a Title/Heading or Body paragraph:
                # 1. Explicit [P0]: Passage title or flush-left line (FirstLineIndent = 0)
                # 2. Explicit [P1]: Passage body paragraph (FirstLineIndent = 0.75cm, Justified)
                # 3. Fallback inference if untagged:
                if sub == "P0":
                    is_p0_title = True
                elif sub in ("P1", "P2"):
                    is_p0_title = False
                else:
                    is_p0_title = (len(raw_c) <= 70 and (raw_c.startswith('*') or raw_c.startswith('**') or raw_c.startswith('_')))

                space_before_quote = 14 if (self.last_rendered_tag == "BOX") else (6 if is_p0_title else 0)

                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.RightIndent = 0
                sel.ParagraphFormat.FirstLineIndent = 0 if is_p0_title else cm_to_pt(0.75)
                sel.ParagraphFormat.SpaceBefore = space_before_quote
                sel.ParagraphFormat.SpaceAfter = 3 if is_p0_title else 4
                sel.ParagraphFormat.KeepWithNext = is_p0_title
                # Alignment: Left (0) for P0 titles/headings, Justified (3) for P1 body paragraphs
                sel.ParagraphFormat.Alignment = 0 if is_p0_title else 3

                self.write_inline_spans(sel, block.spans, default_italic=False)
                sel.TypeParagraph()
                sel.ParagraphFormat.RightIndent = 0
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.FirstLineIndent = 0
                sel.ParagraphFormat.KeepWithNext = False

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
                sel = word.Selection
                word.ActiveWindow.ScrollIntoView(sel.Range, True)
            except Exception:
                pass

            self.last_rendered_tag = tag
            idx_block += 1

        if is_root:
            self.apply_page_numbers(doc)
