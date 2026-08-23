import os
import re
import math
from typing import List, Dict, Any, Optional, Tuple
from uln_parser import ULNBlock, InlineSpan, PicInfo, parse_pic_tag, parse_inline_spans
from renderer_utils import (
    cm_to_pt,
    pt_to_cm,
    parse_color_to_rgb_int,
    extract_question_prefix_and_body,
    split_line_into_option_items,
    get_gdi_text_measurer
)

class RendererBlocksMixin:
    """
    Mixin containing specialized block-level renderers for MS Word documents:
    - Tables (bordered & borderless)
    - Box Shapes (Formulas, Callouts, Word Banks)
    - Multiple-Choice Option Grids ([OPT])
    - Picture Grids ([PIC_GRID]) & Diagrams ([PIC])
    - Auto-numbered Exercise Containers ([NUM])
    """

    def render_pic(self, sel, doc, pic: PicInfo):
        """Renders an image file from user queue in order, or falls back to 'test pic/' folder."""
        target_path = self.get_next_image_path(pic)

        if target_path and os.path.exists(target_path):
            try:
                shape = sel.InlineShapes.AddPicture(FileName=os.path.abspath(target_path))
                # Enforce UNIFORM identical width and height across all pictures in the exercise
                try:
                    shape.LockAspectRatio = 0
                except Exception:
                    pass

                is_in_table = False
                try:
                    is_in_table = bool(sel.Information(12))  # wdWithInTable = 12
                except Exception:
                    pass

                if getattr(self, 'current_tab2_pic_width_cm', None):
                    w_cm = self.current_tab2_pic_width_cm
                    h_cm = getattr(self, 'current_tab2_pic_height_cm', w_cm * 0.72)
                    shape.Width = cm_to_pt(w_cm)
                    shape.Height = cm_to_pt(h_cm)
                elif is_in_table or pic.size == "small":
                    shape.Width = cm_to_pt(3.8)
                    shape.Height = cm_to_pt(2.7)
                elif pic.size == "large":
                    shape.Width = cm_to_pt(9.0)
                    shape.Height = cm_to_pt(6.0)
                else:
                    # Standalone medium picture
                    shape.Width = cm_to_pt(5.0)
                    shape.Height = cm_to_pt(3.5)
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
        if b.cols:
            b.cols = [process_text_num(c) for c in b.cols]
        if b.spans:
            for span in b.spans:
                span.text = process_text_num(span.text)
        if b.col1_spans:
            for span in b.col1_spans:
                span.text = process_text_num(span.text)
        if b.col2_spans:
            for span in b.col2_spans:
                span.text = process_text_num(span.text)
        if b.cols_spans:
            for spans_list in b.cols_spans:
                for span in spans_list:
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
                body_t = m_let.group(2) if m_let else clean_item
                clean_body = re.sub(r'\[(.*?)\]\{(?:u|b|i|[a-zA-Z0-9#]+)\}', r'\1', body_t)
                clean_body = re.sub(r'\[.*?\]|\{.*?\}|\*|_', '', clean_body).strip()
                item_str = f"{m_let.group(1)} {clean_body}" if m_let else f"A. {clean_body}"
                all_items.append(item_str)

        if all_items:
            left_indent_cm = 0.5
            cols = self.calculate_optimal_option_cols(all_items, left_indent_cm, printable_width_cm)
            clean_items = [self.strip_markup_for_measurement(i) for i in all_items]
            measurer = get_gdi_text_measurer()
            max_w_pt = max(measurer.measure_text_pt(ci, font_name=self.font_name, font_size_pt=self.font_size, is_bold=False) for ci in clean_items) if clean_items else 0.0
            return cols, max_w_pt
        return None, None

    def render_num_container(self, sel, doc, word, block: ULNBlock, printable_width_cm: float):
        """
        Renders auto-numbered container [NUM] ... [/NUM].
        Flags the first question in this section to start a new independent list.
        Pre-computes uniform option alignment across all questions in the exercise.
        """
        if not block.children:
            return

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
        based on exact physical GDI typographical width with Word metrics (1.15x) so text wrapping NEVER occurs.
        """
        N = len(items)
        if N <= 1:
            return 1

        remaining_width_cm = max(5.0, printable_width_cm - left_indent_cm)
        remaining_width_pt = cm_to_pt(remaining_width_cm)

        # Measure exact physical width in points of each option item with Word typographical scaling (1.15x)
        clean_items = [self.strip_markup_for_measurement(item) for item in items]
        measurer = get_gdi_text_measurer()
        item_widths_pt = [measurer.measure_text_pt(ci, font_name=self.font_name, font_size_pt=self.font_size, is_bold=False) * 1.15 for ci in clean_items]
        max_item_w_pt = max(item_widths_pt) if item_widths_pt else 0.0

        # Minimum safety distance between columns (strictly >= 5.0 mm = 0.50 cm)
        col_gap_pt = cm_to_pt(0.50)

        # Candidate column counts to evaluate based on total items
        if N >= 4:
            candidate_cols = [4, 2]
        elif N == 3:
            candidate_cols = [3]
        elif N == 2:
            candidate_cols = [2]
        else:
            candidate_cols = []

        for c in candidate_cols:
            col_slot_w_pt = (remaining_width_pt - ((c - 1) * col_gap_pt)) / c
            # If the longest item fits comfortably inside the column slot
            if max_item_w_pt <= col_slot_w_pt:
                return c

        return 1

    def render_opt(self, sel, doc, word, block: ULNBlock, printable_width_cm: float):
        """
        Renders dedicated multiple-choice option container [OPT] ... [/OPT].
        Automatically formats option letters (A., B., C., D.) as bold, and calculates optimal column count
        (1, 2, 3, or 4 columns) based on exact GDI physical text width so text wrapping NEVER occurs.
        """
        raw_text = block.content.strip()
        if not raw_text:
            return

        if '|' in raw_text:
            items = [x.strip() for x in raw_text.split('|') if x.strip()]
        elif '\n' in raw_text:
            items = [x.strip() for x in raw_text.split('\n') if x.strip()]
        else:
            items = split_line_into_option_items(raw_text)

        if not items:
            return

        # Check if first item has a question number (Pronunciation/Stress question with options-only)
        q_num_extracted = None
        has_standalone_q_num = False
        extracted_pref, extracted_delim, extracted_num, clean_item_0 = extract_question_prefix_and_body(items[0])
        if extracted_num is not None:
            has_standalone_q_num = True
            q_num_extracted = extracted_num
            items[0] = clean_item_0
        else:
            q_match = re.match(r'^\s*(?:#?(\d+)[\.\)]|Question\s+#?(\d+)[\.\)]?|Câu\s+#?(\d+)[\.\)]?)\s*(.*)$', items[0], re.IGNORECASE)
            if q_match:
                has_standalone_q_num = True
                q_num_extracted = q_match.group(1) or q_match.group(2) or q_match.group(3)
                items[0] = q_match.group(4).strip() if q_match.group(4) else items[0]

        normalized_items = []
        for idx_item, item in enumerate(items):
            m_let = re.match(r'^\s*(?:(?:\*\*|\*|\[|\(?)*([a-zA-Z][\.\)])(?:\*\*|\*|\]|\}|\{u\}|\))*)\s*(.*)$', item)
            if m_let:
                let_part = m_let.group(1).rstrip('.)')
                body_part = m_let.group(2).strip()
                normalized_items.append((let_part, body_part))
            else:
                let_part = chr(65 + idx_item) if idx_item < 26 else str(idx_item + 1)
                normalized_items.append((let_part, item))

        left_indent_cm = 0.0 if has_standalone_q_num else 0.5

        formatted_item_strings = [f"{let}. {self.strip_markup_for_measurement(body)}" for let, body in normalized_items]
        local_cols = self.calculate_optimal_option_cols(formatted_item_strings, left_indent_cm, printable_width_cm)
        measurer = get_gdi_text_measurer()
        local_max_w_pt = max(measurer.measure_text_pt(s, font_name=self.font_name, font_size_pt=self.font_size) for s in formatted_item_strings) if formatted_item_strings else 0.0

        # Each question independently uses its own optimal column count based on its actual option lengths
        num_cols = local_cols
        max_item_len = local_max_w_pt

        if has_standalone_q_num:
            num_fmt = self.get_effective_number_format(extracted_pref, extracted_delim)
            self.apply_native_numbered_list(word, sel, q_num=q_num_extracted, number_format=num_fmt)
        else:
            try:
                sel.Range.ListFormat.RemoveNumbers()
            except Exception:
                pass
            sel.ParagraphFormat.LeftIndent = cm_to_pt(left_indent_cm)
            sel.ParagraphFormat.FirstLineIndent = 0

        sel.ParagraphFormat.SpaceBefore = 3
        sel.ParagraphFormat.SpaceAfter = 3
        sel.ParagraphFormat.LineSpacingRule = 0
        sel.ParagraphFormat.Alignment = 0
        sel.ParagraphFormat.KeepWithNext = False

        opt_color_int = parse_color_to_rgb_int(getattr(self, 'opt_color', '#000000'))

        if num_cols >= 2 and len(normalized_items) >= 2:
            self.setup_tab_stops(sel, num_cols, left_indent_cm, printable_width_cm, max_item_len=max_item_len)

            for idx_opt, (opt_letter, opt_body) in enumerate(normalized_items):
                col_idx = idx_opt % num_cols

                if idx_opt > 0 and col_idx == 0:
                    sel.TypeParagraph()
                    if has_standalone_q_num:
                        try:
                            sel.Range.ListFormat.RemoveNumbers()
                        except Exception:
                            pass
                        sel.ParagraphFormat.LeftIndent = cm_to_pt(0.5)
                        sel.ParagraphFormat.FirstLineIndent = 0

                sel.Font.Name = self.font_name
                sel.Font.Size = self.font_size
                sel.Font.Bold = 1
                sel.Font.Italic = 0
                sel.Font.Underline = 0
                if opt_color_int is not None:
                    try:
                        sel.Font.Color = opt_color_int
                    except Exception:
                        pass
                else:
                    sel.Font.Color = 0

                sel.TypeText(f"{opt_letter}. ")

                sel.Font.Bold = 0
                sel.Font.Color = 0

                spans = parse_inline_spans(opt_body)
                self.write_inline_spans(sel, spans)
                try:
                    sel.Font.Underline = 0
                except Exception:
                    pass

                if col_idx < num_cols - 1 and idx_opt < len(normalized_items) - 1:
                    try:
                        sel.Font.Underline = 0
                    except Exception:
                        pass
                    sel.TypeText("\t")

            try:
                sel.Font.Underline = 0
            except Exception:
                pass
            sel.TypeParagraph()
        else:
            # 1-column layout
            for idx_opt, (opt_letter, opt_body) in enumerate(normalized_items):
                if idx_opt > 0:
                    if has_standalone_q_num:
                        try:
                            sel.Range.ListFormat.RemoveNumbers()
                        except Exception:
                            pass
                        sel.ParagraphFormat.LeftIndent = cm_to_pt(0.5)
                        sel.ParagraphFormat.FirstLineIndent = 0

                sel.Font.Name = self.font_name
                sel.Font.Size = self.font_size
                sel.Font.Bold = 1
                sel.Font.Italic = 0
                sel.Font.Underline = 0
                if opt_color_int is not None:
                    try:
                        sel.Font.Color = opt_color_int
                    except Exception:
                        pass
                else:
                    sel.Font.Color = 0

                sel.TypeText(f"{opt_letter}. ")

                sel.Font.Bold = 0
                sel.Font.Color = 0

                spans = parse_inline_spans(opt_body)
                self.write_inline_spans(sel, spans)
                try:
                    sel.Font.Underline = 0
                except Exception:
                    pass
                sel.TypeParagraph()

        sel.ParagraphFormat.LeftIndent = 0
        sel.ParagraphFormat.FirstLineIndent = 0
        sel.ParagraphFormat.TabStops.ClearAll()
        self.last_rendered_tag = "OPT"

    def measure_text_width_pt(self, doc, text: str, font_name: str = "Times New Roman", font_size: float = 12.0, is_bold: bool = False, is_italic: bool = False) -> float:
        """Accurately measures physical text width in points using Windows GDI typography engine with caching."""
        if not text:
            return 0.0
        return get_gdi_text_measurer().measure_text_pt(text, font_name=font_name, font_size_pt=font_size, is_bold=is_bold, is_italic=is_italic)

    def strip_markup_for_measurement(self, text: str) -> str:
        """Strips inline ULN tags, bold/italic markers, and answer blanks for clean physical text width measurement."""
        if not text:
            return ""
        # 1. Replace <blank> / [BLANK] with 11 underscores matching write_inline_spans
        t = re.sub(r'<(?:blank|BLANK)>|\[(?:blank|BLANK)\]', '___________', text)
        # 2. Extract inner text from [text]{tag}
        t = re.sub(r'\[(.*?)\]\{(?:u|b|i|[a-zA-Z0-9#:,]+)\}', r'\1', text)
        # 3. Strip [ins], [/ins] and [PIC: ...]
        t = re.sub(r'\[\/?(?:ins|INS)\]', '', t)
        t = re.sub(r'\[PIC(?::.*?)?\]', '', t, flags=re.IGNORECASE)
        # 4. Strip markdown bold/italic asterisks while preserving inner text
        t = re.sub(r'\*\*(.*?)\*\*', r'\1', t)
        t = re.sub(r'\*(.*?)\*', r'\1', t)
        return t.strip()

    def calculate_item_visual_lines(self, doc, text: str, avail_width_pt: float, is_bold: bool = True) -> int:
        """
        Calculates the exact visual wrapped line count of a text string when rendered
        within a bounded horizontal width in MS Word.
        """
        clean = self.strip_markup_for_measurement(text)
        if not clean:
            return 1
        words = clean.split()
        if not words:
            return 1
        
        space_w = self.measure_text_width_pt(doc, " ", self.font_name, self.font_size, is_bold=is_bold)
        curr_line_w = 0.0
        lines_count = 1
        
        for w in words:
            w_pt = self.measure_text_width_pt(doc, w, self.font_name, self.font_size, is_bold=is_bold) * 1.15
            if curr_line_w == 0.0:
                curr_line_w = w_pt
            elif curr_line_w + space_w + w_pt <= avail_width_pt:
                curr_line_w += space_w + w_pt
            else:
                lines_count += 1
                curr_line_w = w_pt
                
        return max(1, lines_count)

    def optimize_word_bank_layout(self, doc, words: List[str], printable_width_pt: float) -> Tuple[int, List[List[str]], float, List[float]]:
        """
        Optimally arranges Word Bank items across columns and rows to produce the most
        compact, balanced, and aesthetically pleasing box width and height.
        Evaluates column-major alphabetical, length-balanced bin-packing, and row-major permutations.
        """
        N = len(words)
        if N == 0:
            return 1, [[]], 50.0, [0.0]

        clean_words = [self.strip_markup_for_measurement(w) for w in words]
        item_widths_pt = [self.measure_text_width_pt(doc, cw, self.font_name, self.font_size, is_bold=True) * 1.15 for cw in clean_words]
        max_item_w_pt = max(item_widths_pt) if item_widths_pt else 45.0

        pad_horiz_pt = cm_to_pt(0.20)      # 2.0 mm padding
        extra_buffer_pt = cm_to_pt(0.20)   # 2.0 mm corner buffer
        gap_pt = cm_to_pt(0.80)            # 8.0 mm inter-column gap

        # If items are exceptionally long sentences (>72% page width), use single column sized to median
        if max_item_w_pt >= (printable_width_pt * 0.72):
            sorted_w = sorted(item_widths_pt)
            median_w = sorted_w[len(sorted_w) // 2]
            target_w = max(median_w, sorted_w[int(len(sorted_w) * 0.55)])
            box_w = min(printable_width_pt * 0.76, max(printable_width_pt * 0.50, target_w + (2 * pad_horiz_pt) + extra_buffer_pt))
            return 1, [[w] for w in words], box_w, [0.0]

        items_with_w = list(zip(words, item_widths_pt))

        best_c = 1
        best_grid = [[w] for w in words]
        best_box_w = min(printable_width_pt, max(item_widths_pt) + (2 * pad_horiz_pt) + extra_buffer_pt)
        best_tabs = [0.0]
        best_score = float('inf')

        max_c = min(N, 6)
        for c in range(max_c, 0, -1):
            num_rows = math.ceil(N / c)

            arrangements = []

            # 1. Alphabetical Column-Major (standard dictionary flow: Col 1 top-to-bottom, then Col 2...)
            alpha_items = sorted(items_with_w, key=lambda x: self.strip_markup_for_measurement(x[0]).lower())
            grid_col_major = [[] for _ in range(num_rows)]
            for i, item in enumerate(alpha_items):
                col_idx = i // num_rows
                row_idx = i % num_rows
                if row_idx < num_rows and col_idx < c:
                    grid_col_major[row_idx].append(item)
            grid_col_major = [r for r in grid_col_major if r]
            arrangements.append(grid_col_major)

            # 2. Length-Balanced (Greedy Bin-Packing into lowest-width columns)
            by_len = sorted(items_with_w, key=lambda x: x[1], reverse=True)
            col_buckets = [[] for _ in range(c)]
            for it in by_len:
                valid_buckets = [b_idx for b_idx in range(c) if len(col_buckets[b_idx]) < num_rows]
                if not valid_buckets:
                    break
                chosen_b = min(valid_buckets, key=lambda b_idx: sum(x[1] for x in col_buckets[b_idx]))
                col_buckets[chosen_b].append(it)
            grid_balanced = [[] for _ in range(num_rows)]
            for b_idx in range(c):
                for r_idx, it in enumerate(col_buckets[b_idx]):
                    grid_balanced[r_idx].append(it)
            grid_balanced = [r for r in grid_balanced if r]
            arrangements.append(grid_balanced)

            # 3. Alphabetical Row-Major (A, B, C, D left-to-right)
            grid_row_major = []
            for r in range(num_rows):
                chunk = alpha_items[r*c : (r+1)*c]
                if chunk:
                    grid_row_major.append(chunk)
            arrangements.append(grid_row_major)

            # 4. Original input order
            grid_orig = []
            for r in range(num_rows):
                chunk = items_with_w[r*c : (r+1)*c]
                if chunk:
                    grid_orig.append(chunk)
            arrangements.append(grid_orig)

            for grid in arrangements:
                if not grid or not grid[0]:
                    continue
                actual_c = max(len(row) for row in grid)
                if actual_c == 0:
                    continue
                col_max_w = []
                for col_idx in range(actual_c):
                    col_items = [row[col_idx][1] for row in grid if col_idx < len(row)]
                    col_max_w.append(max(col_items) if col_items else 40.0)

                needed_w = sum(col_max_w) + ((actual_c - 1) * gap_pt) + (2 * pad_horiz_pt) + extra_buffer_pt
                if needed_w <= printable_width_pt:
                    # Score: heavily favor fewer rows (R), then favor compact balanced width
                    score = num_rows * 1000 + needed_w
                    if score < best_score:
                        best_score = score
                        best_c = actual_c
                        best_grid = [[x[0] for x in row] for row in grid]

                        tabs = [0.0]
                        for ci in range(actual_c - 1):
                            tabs.append(tabs[-1] + col_max_w[ci] + gap_pt)
                        best_tabs = tabs
                        best_box_w = min(printable_width_pt, tabs[-1] + col_max_w[-1] + (2 * pad_horiz_pt) + extra_buffer_pt)

        return best_c, best_grid, best_box_w, best_tabs

    def render_box_shape(self, sel, doc, word, block: ULNBlock, printable_width_cm: float, idx_block: int = 0, blocks: List[ULNBlock] = None):
        """
        Renders Word Bank / Callout Box / Formula Box inside a MS Word Rounded Rectangle Shape TextFrame.
        - Sets KeepWithNext ONLY if immediately adjacent to an [INS] block.
        - Word Bank automatically optimizes item layout and column assignments for minimal height and balanced width.
        """
        printable_width_pt = cm_to_pt(printable_width_cm)
        raw_content = block.content.strip() if block.content else ""
        if not raw_content:
            return

        is_word_bank = ('|' in raw_content) or (block.tag == "WORDBANK") or (block.tag.endswith(":bank"))

        is_adjacent_to_ins = False
        if blocks and 0 <= idx_block < len(blocks):
            if idx_block > 0 and blocks[idx_block - 1].tag == "INS":
                is_adjacent_to_ins = True
            elif idx_block + 1 < len(blocks) and blocks[idx_block + 1].tag == "INS":
                is_adjacent_to_ins = True
        elif self.last_rendered_tag == "INS":
            is_adjacent_to_ins = True

        p_anchor = doc.Range(sel.Range.Start, sel.Range.Start)
        try:
            p_anchor.ParagraphFormat.SpaceBefore = 14.0
            p_anchor.ParagraphFormat.SpaceAfter = 14.0
            p_anchor.ParagraphFormat.KeepWithNext = is_adjacent_to_ins
        except Exception:
            pass

        if not is_word_bank:
            # PATHWAY A: FORMULA / CALLOUT / TEXT BOX
            lines = [l.strip() for l in raw_content.split('\n') if l.strip()]
            if not lines:
                return

            # Measure exact physical rendered width of clean lines without markup noise
            clean_lines = [self.strip_markup_for_measurement(l) for l in lines]
            line_widths_pt = [self.measure_text_width_pt(doc, cl, self.font_name, self.font_size, is_bold=True) for cl in clean_lines]
            max_line_w_pt = max(line_widths_pt) if line_widths_pt else 50.0

            pad_left_pt = cm_to_pt(0.20)   # Exactly 2.0 mm left margin
            pad_right_pt = cm_to_pt(0.20)  # Exactly 2.0 mm right margin
            pad_top_pt = 0.0               # 0.0 pt top margin
            pad_bottom_pt = 0.0            # 0.0 pt bottom margin
            extra_buffer_pt = cm_to_pt(0.20) # 2.0 mm corner clearance buffer

            total_pad_pt = pad_left_pt + pad_right_pt + extra_buffer_pt

            if (max_line_w_pt + total_pad_pt) <= (printable_width_pt * 0.72):
                box_width_pt = max(80.0, max_line_w_pt + total_pad_pt)
                is_full_width = False
            else:
                # Multi-line text boxes with wrapping lines: size proportionally to median line width
                sorted_w = sorted(line_widths_pt)
                median_w = sorted_w[len(sorted_w) // 2]
                target_w = max(median_w, sorted_w[int(len(sorted_w) * 0.55)])
                target_content_w = target_w + total_pad_pt
                box_width_pt = min(printable_width_pt * 0.76, max(printable_width_pt * 0.50, target_content_w))
                is_full_width = False

            left_offset_pt = max(0.0, (printable_width_pt - box_width_pt) / 2.0)

            # Vertical Height: (Total Visual Lines * Font Line Height) + Space Between Lines + Descender clearance buffer
            num_lines = len(lines)
            exact_line_h_pt = self.font_size * 1.28  # Standard single line height
            space_between_pt = 2.0
            descender_clearance_pt = max(6.0, self.font_size * 0.40)  # Prevents descenders (y, g, p, q, j) from touching bottom border
            avail_inner_w = box_width_pt - pad_left_pt - pad_right_pt - extra_buffer_pt
            total_visual_lines = sum(self.calculate_item_visual_lines(doc, l, avail_inner_w, is_bold=True) for l in lines)
            box_height_pt = (total_visual_lines * exact_line_h_pt) + ((num_lines - 1) * space_between_pt) + descender_clearance_pt + 4.0

            try:
                shape = doc.Shapes.AddShape(
                    5,  # msoShapeRoundedRectangle = 5
                    0,
                    0,
                    box_width_pt,
                    box_height_pt,
                    Anchor=p_anchor
                )
                shape.RelativeHorizontalPosition = 0
                shape.RelativeVerticalPosition = 2
                shape.Left = left_offset_pt
                shape.Top = 0
                shape.WrapFormat.Type = 7
                shape.WrapFormat.DistanceTop = 12.0
                shape.WrapFormat.DistanceBottom = 12.0

                tf = shape.TextFrame
                tf.MarginTop = 0
                tf.MarginBottom = 0
                tf.MarginLeft = pad_left_pt
                tf.MarginRight = pad_right_pt
                try:
                    tf.WordWrap = -1 if (is_full_width or total_visual_lines > num_lines) else 0
                except Exception:
                    pass

                try:
                    tf.AutoSize = False
                except Exception:
                    pass

                shape.Fill.Visible = False
                shape.Line.Weight = 1.0
                shape.Line.ForeColor.RGB = 0

                tf.TextRange.Select()
                box_sel = word.Selection
                box_sel.Font.Name = self.font_name
                box_sel.Font.Size = self.font_size
                box_sel.Font.Bold = 1
                box_sel.Font.Color = 0

                box_sel.ParagraphFormat.SpaceBefore = 0
                box_sel.ParagraphFormat.SpaceAfter = 0
                box_sel.ParagraphFormat.LineSpacingRule = 0
                box_sel.ParagraphFormat.Alignment = 0  # Left-aligned so longest line reaches exact right margin
                box_sel.ParagraphFormat.TabStops.ClearAll()

                for idx_line, line_str in enumerate(lines):
                    if idx_line > 0:
                        box_sel.ParagraphFormat.SpaceBefore = 2.0
                    box_sel.ParagraphFormat.SpaceAfter = 0

                    line_spans = parse_inline_spans(line_str, default_bold=True)
                    self.write_inline_spans(box_sel, line_spans)
                    box_sel.Font.Color = 0
                    if idx_line < len(lines) - 1:
                        box_sel.TypeParagraph()

                try:
                    shape.ConvertToInlineShape()
                except Exception:
                    pass

            except Exception as e:
                print(f"[ULNRenderer] Warning creating Formula/Callout TextFrame box shape: {e}")

        else:
            # PATHWAY B: WORD BANK / PIPE-SEPARATED CHOICES (OPTIMAL COLUMN & ARRANGEMENT SIZING)
            words = [w.strip() for w in raw_content.split('|') if w.strip()]
            if not words:
                return

            cols, lines_bank, box_width_pt, tab_stops_pt = self.optimize_word_bank_layout(doc, words, printable_width_pt)
            pad_horiz_pt = cm_to_pt(0.20)  # Exactly 2.0 mm padding
            extra_buffer_pt = cm_to_pt(0.20)
            left_offset_pt = max(0.0, (printable_width_pt - box_width_pt) / 2.0)

            # Calculate total visual lines accounting for wrapped text lines in every row
            avail_inner_w = box_width_pt - (2 * pad_horiz_pt) - extra_buffer_pt
            total_visual_lines = 0
            if cols == 1:
                for chunk in lines_bank:
                    item_text = chunk[0] if chunk else ""
                    v_lines = self.calculate_item_visual_lines(doc, item_text, avail_inner_w, is_bold=True)
                    total_visual_lines += v_lines
            else:
                # Multi-column: determine column widths from tab_stops
                num_cols = len(tab_stops_pt)
                col_widths = []
                for ci in range(num_cols):
                    if ci + 1 < num_cols:
                        col_widths.append(tab_stops_pt[ci + 1] - tab_stops_pt[ci] - cm_to_pt(0.40))
                    else:
                        col_widths.append(box_width_pt - tab_stops_pt[ci] - (2 * pad_horiz_pt))
                for chunk in lines_bank:
                    row_v_lines = 1
                    for ci, item_text in enumerate(chunk):
                        c_w = col_widths[ci] if ci < len(col_widths) else avail_inner_w
                        v_l = self.calculate_item_visual_lines(doc, item_text, c_w, is_bold=True)
                        if v_l > row_v_lines:
                            row_v_lines = v_l
                    total_visual_lines += row_v_lines

            num_rows = len(lines_bank)
            exact_line_h_pt = self.font_size * 1.28
            space_between_pt = 2.0
            descender_clearance_pt = max(6.0, self.font_size * 0.40)  # Prevents descenders (y, g, p, q, j) from touching bottom border
            box_height_pt = (total_visual_lines * exact_line_h_pt) + ((num_rows - 1) * space_between_pt) + descender_clearance_pt + 4.0

            try:
                shape = doc.Shapes.AddShape(
                    5,
                    0,
                    0,
                    box_width_pt,
                    box_height_pt,
                    Anchor=p_anchor
                )
                shape.RelativeHorizontalPosition = 0
                shape.RelativeVerticalPosition = 2
                shape.Left = left_offset_pt
                shape.Top = 0
                shape.WrapFormat.Type = 7
                shape.WrapFormat.DistanceTop = 12.0
                shape.WrapFormat.DistanceBottom = 12.0

                tf = shape.TextFrame
                tf.MarginTop = 0
                tf.MarginBottom = 0
                tf.MarginLeft = pad_horiz_pt
                tf.MarginRight = pad_horiz_pt
                try:
                    tf.WordWrap = -1 if (cols == 1 or total_visual_lines > num_rows) else 0
                except Exception:
                    pass
                try:
                    tf.AutoSize = False
                except Exception:
                    pass

                shape.Fill.Visible = False
                shape.Line.Weight = 1.0
                shape.Line.ForeColor.RGB = 0

                tf.TextRange.Select()
                box_sel = word.Selection
                box_sel.Font.Name = self.font_name
                box_sel.Font.Size = self.font_size
                box_sel.Font.Bold = 1
                box_sel.Font.Color = 0

                box_sel.ParagraphFormat.SpaceBefore = 0
                box_sel.ParagraphFormat.SpaceAfter = 0
                box_sel.ParagraphFormat.LineSpacingRule = 0
                box_sel.ParagraphFormat.Alignment = 0
                box_sel.ParagraphFormat.TabStops.ClearAll()

                for t_pt in tab_stops_pt[1:]:
                    box_sel.ParagraphFormat.TabStops.Add(Position=t_pt, Alignment=0)

                for idx_line, chunk in enumerate(lines_bank):
                    if idx_line > 0:
                        box_sel.ParagraphFormat.SpaceBefore = 2.0
                    box_sel.ParagraphFormat.SpaceAfter = 0

                    for idx_w, word_txt in enumerate(chunk):
                        w_spans = parse_inline_spans(word_txt, default_bold=True)
                        self.write_inline_spans(box_sel, w_spans)
                        box_sel.Font.Color = 0
                        if idx_w < len(chunk) - 1:
                            box_sel.TypeText("\t")

                    if idx_line < len(lines_bank) - 1:
                        box_sel.TypeParagraph()

                try:
                    shape.ConvertToInlineShape()
                except Exception:
                    pass

            except Exception as e:
                print(f"[ULNRenderer] Warning creating Word Bank TextFrame box shape: {e}")

        # Common epilogue
        try:
            end_range = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
            end_range.Select()
            sel.ParagraphFormat.LeftIndent = 0
            sel.ParagraphFormat.RightIndent = 0
            sel.ParagraphFormat.Alignment = 1
            sel.ParagraphFormat.SpaceBefore = 14.0
            sel.ParagraphFormat.SpaceAfter = 14.0
            sel.TypeParagraph()
            sel.ParagraphFormat.SpaceBefore = 0
            sel.ParagraphFormat.SpaceAfter = 4
            sel.ParagraphFormat.Alignment = 0
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

        total_pics = sum(
            1 for row in tdata.rows for cell in row.cells
            if ("[PIC" in cell.content.upper() or parse_pic_tag(cell.content) is not None)
        )
        has_pic = (total_pics > 0)

        # Only use floating side-diagram layout if it is a single MCQ question with a side diagram
        is_single_diagram_mcq = (
            tdata.borderless
            and total_pics == 1
            and len(tdata.rows) >= 3
            and any(re.match(r'^\s*\*?\*?[A-Da-d][\.\)]', cell.content) for row in tdata.rows for cell in row.cells)
        )

        if is_single_diagram_mcq:
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

            for idx_r, txt_line in enumerate(text_rows):
                txt_line = re.sub(r'#(\d+)', r'\1', txt_line)

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

                spans = parse_inline_spans(txt_line)

                sel.ParagraphFormat.SpaceBefore = 14 if (idx_r == 0 and self.last_rendered_tag == "BOX") else 3
                sel.ParagraphFormat.SpaceAfter = 3
                sel.ParagraphFormat.LineSpacingRule = 0
                sel.ParagraphFormat.Alignment = 0

                if idx_r == 1 or (is_opt_line and options_anchor_range is None):
                    try:
                        options_anchor_range = sel.Range.Duplicate
                        opt_start_top_pt = options_anchor_range.Information(6)
                    except Exception:
                        pass

                self.write_inline_spans(sel, spans)
                sel.TypeParagraph()

            sel.ParagraphFormat.LeftIndent = 0

            opt_end_top_pt = 0.0
            try:
                opt_end_top_pt = sel.Range.Information(6)
            except Exception:
                pass

            if has_pic and pic_info_found:
                pic_path = self.get_next_image_path(pic_info_found)
                if pic_path and os.path.exists(pic_path):
                    try:
                        total_opt_height_pt = max(60.0, opt_end_top_pt - opt_start_top_pt)
                        img_h_pt = max(60.0, min(140.0, total_opt_height_pt - 6.0))
                        img_w_pt = max(75.0, min(160.0, cm_to_pt(5.5)))

                        anchor_rng = options_anchor_range if options_anchor_range else sel.Range
                        shp = doc.Shapes.AddPicture(
                            FileName=os.path.abspath(pic_path),
                            LinkToFile=False,
                            SaveWithDocument=True,
                            Anchor=anchor_rng
                        )
                        shp.RelativeHorizontalPosition = 0
                        shp.RelativeVerticalPosition = 2
                        shp.Left = cm_to_pt(printable_width_cm) - img_w_pt
                        shp.Top = 0
                        shp.Width = img_w_pt
                        shp.Height = img_h_pt
                        shp.WrapFormat.Type = 1
                    except Exception as e:
                        print(f"[ULNRenderer] Warning adding borderless diagram picture: {e}")

            sel.ParagraphFormat.LeftIndent = 0
            sel.ParagraphFormat.FirstLineIndent = 0
            self.last_rendered_tag = "TABLE"
            return

        if tdata.borderless and not has_pic:
            # Render borderless multi-column items using native Paragraph Tab Stops (NO Word Table object)
            col_w_cm = printable_width_cm / max(1, num_cols)
            
            for row in tdata.rows:
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.FirstLineIndent = 0
                sel.ParagraphFormat.SpaceBefore = 2
                sel.ParagraphFormat.SpaceAfter = 2
                sel.ParagraphFormat.LineSpacingRule = 0
                sel.ParagraphFormat.Alignment = 0
                sel.ParagraphFormat.TabStops.ClearAll()

                for c in range(1, num_cols):
                    sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(col_w_cm * c), Alignment=0)

                try:
                    sel.Range.ListFormat.RemoveNumbers()
                except Exception:
                    pass

                for idx_cell, cell in enumerate(row.cells):
                    if idx_cell > 0:
                        sel.TypeText("\t")
                    if cell.spans:
                        self.write_inline_spans(sel, cell.spans, default_bold=cell.is_header)
                    else:
                        sel.Font.Bold = 1 if cell.is_header else 0
                        sel.TypeText(cell.content.strip())

                sel.TypeParagraph()
                sel.ParagraphFormat.TabStops.ClearAll()

            self.last_rendered_tag = "TABLE"
            return

        # Native Table Layout (Bordered Grid)
        p_table_anchor = doc.Range(sel.Range.Start, sel.Range.Start)
        try:
            p_table_anchor.ParagraphFormat.SpaceBefore = 14.0 if (self.last_rendered_tag == "BOX") else 8.0
            p_table_anchor.ParagraphFormat.SpaceAfter = 14.0
        except Exception:
            pass

        tbl = doc.Tables.Add(Range=p_table_anchor, NumRows=num_rows, NumColumns=num_cols)
        tbl.AllowAutoFit = False

        if tdata.borderless:
            tbl.Borders.Enable = False
        else:
            try:
                tbl.Borders.InsideLineStyle = 1
                tbl.Borders.InsideLineWidth = 8
                tbl.Borders.InsideColor = 0
                tbl.Borders.OutsideLineStyle = 1
                tbl.Borders.OutsideLineWidth = 8
                tbl.Borders.OutsideColor = 0
            except Exception:
                pass

        col_w_pt = cm_to_pt(printable_width_cm) / max(1, num_cols)
        for col_idx in range(1, num_cols + 1):
            try:
                tbl.Columns(col_idx).Width = col_w_pt
            except Exception:
                pass

        try:
            tbl.TopPadding = cm_to_pt(0.2)
            tbl.BottomPadding = cm_to_pt(0.2)
            tbl.LeftPadding = cm_to_pt(0.2)
            tbl.RightPadding = cm_to_pt(0.2)
        except Exception:
            pass

        try:
            tbl.Range.ListFormat.RemoveNumbers()
        except Exception:
            pass

        for r_idx, row_obj in enumerate(tdata.rows, 1):
            try:
                tbl.Rows(r_idx).AllowBreakAcrossPages = False
            except Exception:
                pass

            for c_idx, cell_obj in enumerate(row_obj.cells, 1):
                cell = tbl.Cell(r_idx, c_idx)
                try:
                    cell.VerticalAlignment = 1  # wdCellAlignVerticalCenter = 1
                except Exception:
                    pass
                cell_range = cell.Range
                try:
                    cell_range.ListFormat.RemoveNumbers()
                except Exception:
                    pass
                cell_range.Font.Name = self.font_name
                cell_range.Font.Size = self.font_size
                cell_range.ParagraphFormat.SpaceBefore = 0
                cell_range.ParagraphFormat.SpaceAfter = 0
                cell_range.ParagraphFormat.LineSpacingRule = 0
                cell_range.ParagraphFormat.Alignment = 0

                cell_txt = cell_obj.content.strip()
                if re.match(r'^\s*(?:_{2,}|<blank>|\[BLANK\])\s*$', cell_txt, re.IGNORECASE):
                    cell_range.Text = ""
                    continue

                if cell_obj.spans:
                    cell_range.Select()
                    cell_sel = doc.Application.Selection
                    try:
                        cell_sel.Range.ListFormat.RemoveNumbers()
                    except Exception:
                        pass
                    self.write_inline_spans(cell_sel, cell_obj.spans, default_bold=cell_obj.is_header)
                else:
                    cell_range.Bold = 1 if cell_obj.is_header else 0
                    cell_range.Text = cell_txt

        try:
            rng_after = doc.Range(tbl.Range.End, tbl.Range.End)
            rng_after.Select()
            sel.ParagraphFormat.LeftIndent = 0
            sel.ParagraphFormat.RightIndent = 0
            sel.ParagraphFormat.FirstLineIndent = 0
            sel.ParagraphFormat.SpaceBefore = 8
            sel.ParagraphFormat.SpaceAfter = 4
            sel.ParagraphFormat.Alignment = 0
            sel.TypeParagraph()
        except Exception:
            pass

        self.last_rendered_tag = "TABLE"

    def render_pic_grid(self, sel, doc, children: List[ULNBlock], printable_width_cm: float):
        """
        Renders 4-column horizontal picture grid [PIC_GRID] using pure paragraph Tab Stops (no Table object).
        Each row has a picture line and an aligned caption line using identical Center Tab Stops.
        """
        if not children:
            return

        N = len(children)
        cols = min(4, N)
        slot_w_cm = printable_width_cm / cols

        for idx_chunk in range(0, N, cols):
            chunk = children[idx_chunk:idx_chunk + cols]
            idx_row = idx_chunk // cols

            # ── 1. Picture Line ──────────────────────────────────────────
            sel.ParagraphFormat.TabStops.ClearAll()
            for c in range(cols):
                center_pos_cm = slot_w_cm * (c + 0.5)
                sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(center_pos_cm), Alignment=1)  # wdAlignTabCenter = 1

            sel.ParagraphFormat.LeftIndent = 0
            sel.ParagraphFormat.RightIndent = 0
            sel.ParagraphFormat.FirstLineIndent = 0
            sel.ParagraphFormat.Alignment = 0
            sel.ParagraphFormat.SpaceBefore = 8 if idx_row == 0 else 4
            sel.ParagraphFormat.SpaceAfter = 2
            sel.ParagraphFormat.LineSpacingRule = 0
            sel.ParagraphFormat.KeepWithNext = True

            for c_idx, child in enumerate(chunk):
                global_idx = idx_chunk + c_idx
                sel.TypeText("\t")

                pic_info = child.pic or parse_pic_tag(child.content) or PicInfo(description=f"Picture {global_idx + 1}", pos="center", size="small")
                target_path = self.get_next_image_path(pic_info)

                if target_path and os.path.exists(target_path):
                    try:
                        col_w_pt = cm_to_pt(slot_w_cm)
                        shp = sel.InlineShapes.AddPicture(FileName=os.path.abspath(target_path))
                        shp.Width = min(col_w_pt - 10.0, cm_to_pt(3.6))
                        shp.Height = cm_to_pt(2.6)
                    except Exception as e:
                        print(f"[ULNRenderer] Warning in pic_grid picture: {e}")
                else:
                    sel.Font.Name = self.font_name
                    sel.Font.Size = 9.0
                    sel.Font.Italic = True
                    sel.Font.Bold = True
                    try:
                        sel.Font.Color = 8421504  # Grey
                    except Exception:
                        pass
                    sel.TypeText(f"[ 🖼️ {global_idx + 1} ]")

            sel.TypeParagraph()

            # ── 2. Caption Line ──────────────────────────────────────────
            sel.ParagraphFormat.TabStops.ClearAll()
            for c in range(cols):
                center_pos_cm = slot_w_cm * (c + 0.5)
                sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(center_pos_cm), Alignment=1)  # wdAlignTabCenter = 1

            sel.ParagraphFormat.LeftIndent = 0
            sel.ParagraphFormat.RightIndent = 0
            sel.ParagraphFormat.FirstLineIndent = 0
            sel.ParagraphFormat.Alignment = 0
            sel.ParagraphFormat.SpaceBefore = 2
            sel.ParagraphFormat.SpaceAfter = 6
            sel.ParagraphFormat.LineSpacingRule = 0
            sel.ParagraphFormat.KeepWithNext = False

            for c_idx, child in enumerate(chunk):
                global_idx = idx_chunk + c_idx
                sel.TypeText("\t")
                sel.Font.Name = self.font_name
                sel.Font.Size = self.font_size
                sel.Font.Bold = 0
                sel.Font.Italic = 0
                sel.Font.Color = 0

                col_w_pt = cm_to_pt(slot_w_cm)
                img_w_pt = min(col_w_pt - 10.0, cm_to_pt(3.6))

                clean_content = re.sub(r'\[PIC(?::[^\]]*)?\]', '', child.content, flags=re.IGNORECASE).strip()
                m_num = re.match(r'^(?:#?(\d+)[\.\)]\s*)?(.*)$', clean_content)
                if m_num:
                    num_part = m_num.group(1) or str(global_idx + 1)
                    body_part = m_num.group(2).strip()
                    if not body_part or "<blank>" in body_part.lower() or "[blank]" in body_part.lower() or "_" in body_part:
                        prefix_w_pt = len(f"{num_part}. ") * (self.font_size * 0.48)
                        char_under_w_pt = max(4.0, self.font_size * 0.44)
                        num_underscores = max(10, int((img_w_pt - prefix_w_pt) / char_under_w_pt))
                        cap_text = f"{num_part}. {'_' * num_underscores}"
                    else:
                        cap_text = f"{num_part}. {body_part}"
                else:
                    prefix_w_pt = len(f"{global_idx + 1}. ") * (self.font_size * 0.48)
                    char_under_w_pt = max(4.0, self.font_size * 0.44)
                    num_underscores = max(10, int((img_w_pt - prefix_w_pt) / char_under_w_pt))
                    cap_text = f"{global_idx + 1}. {'_' * num_underscores}"

                sel.TypeText(cap_text)

            sel.TypeParagraph()

        sel.ParagraphFormat.TabStops.ClearAll()
        sel.ParagraphFormat.LeftIndent = 0
        sel.ParagraphFormat.RightIndent = 0
        sel.ParagraphFormat.FirstLineIndent = 0
        sel.ParagraphFormat.SpaceBefore = 6
        sel.ParagraphFormat.SpaceAfter = 4
        sel.ParagraphFormat.Alignment = 0

        self.last_rendered_tag = "PIC_GRID"

    def render_tab_multi(self, sel, doc, block: ULNBlock, idx_block: int, blocks: List[ULNBlock], printable_width_cm: float):
        """
        Renders 3-column (TAB3), 4-column (TAB4), or N-column (TAB) side-by-side items
        using native Word paragraph tab stops with optimal gap calculations and aligned blanks.
        Auto-determines blank length: 2mm separation from text, min 5mm gap to next column.
        """
        num_cols = len(block.cols) if block.cols else (3 if block.tag == "TAB3" else (4 if block.tag == "TAB4" else 3))

        # 1. Collect full group of consecutive TAB blocks with the same column count
        group_start = idx_block
        while group_start > 0 and blocks[group_start - 1].tag.startswith("TAB") and (len(blocks[group_start - 1].cols) if blocks[group_start - 1].cols else 0) == num_cols:
            group_start -= 1

        tab_group = []
        lookahead = group_start
        while lookahead < len(blocks) and blocks[lookahead].tag.startswith("TAB") and (len(blocks[lookahead].cols) if blocks[lookahead].cols else 0) == num_cols:
            tab_group.append(blocks[lookahead])
            lookahead += 1

        group_first_c1 = tab_group[0].cols[0] if (tab_group and tab_group[0].cols) else (block.cols[0] if block.cols else "")
        base_indent_cm = 0.5 if "P1" in group_first_c1 else (1.0 if "P2" in group_first_c1 else 0.0)

        # 2. Check if columns in this group contain answer blanks
        has_blanks = any(
            re.search(r'[_]{2,}|<(?:blank|BLANK)>|\[(?:blank|BLANK)\]', c)
            for b in tab_group for c in (b.cols or [])
        )

        avail_w_cm = printable_width_cm - base_indent_cm
        even_step_cm = avail_w_cm / float(num_cols)
        min_gap_cm = 0.50  # 5.0 mm minimum gap

        col_starts_cm = [base_indent_cm + (i * even_step_cm) for i in range(num_cols)]

        u_pt = self.measure_text_width_pt(doc, '_', self.font_name, self.font_size, is_bold=False)
        u_width_cm = max(0.20, pt_to_cm(u_pt) * 1.15)

        tab_stops_cm = [col_starts_cm[i] for i in range(1, num_cols)]

        # Apply paragraph formatting
        try:
            sel.Range.ListFormat.RemoveNumbers()
        except Exception:
            pass

        sel.ParagraphFormat.LeftIndent = 0
        sel.ParagraphFormat.FirstLineIndent = 0
        sel.ParagraphFormat.SpaceBefore = 2
        sel.ParagraphFormat.SpaceAfter = 2
        sel.ParagraphFormat.KeepWithNext = False
        sel.ParagraphFormat.PageBreakBefore = False
        sel.ParagraphFormat.TabStops.ClearAll()

        for ts in sorted(tab_stops_cm):
            if ts > 0:
                sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(ts), Alignment=0)

        # Render each column in the current row
        for c_idx, col_raw in enumerate(block.cols):
            if c_idx > 0:
                sel.TypeText("\t")

            clean_col = col_raw.strip()
            # Extract number prefix (e.g. 1. , #7. , 13. )
            m_num = re.match(r'^\s*#?(\d+[\.\)])\s*(.*)$', clean_col)
            num_prefix = m_num.group(1) if m_num else ""
            body_content = m_num.group(2) if m_num else clean_col

            # If this column has an answer blank (e.g. _____ or <blank>), separate word from blank
            if has_blanks and re.search(r'[_]{2,}|<(?:blank|BLANK)>|\[(?:blank|BLANK)\]', body_content):
                word_part = re.sub(r'[_]{2,}|<(?:blank|BLANK)>|\[(?:blank|BLANK)\]', '', body_content).strip()

                # Measure this specific item's width (number + word)
                num_str = f"{num_prefix} " if num_prefix else ""
                full_item_str = f"{num_str}{word_part}".strip()
                w_pt = self.measure_text_width_pt(doc, full_item_str, self.font_name, self.font_size, is_bold=False)
                item_w_cm = pt_to_cm(w_pt) * 1.15

                start_c = col_starts_cm[c_idx]
                next_c = col_starts_cm[c_idx + 1] if (c_idx + 1 < num_cols) else printable_width_cm

                # Blank starts 2mm after word
                blank_start_abs = start_c + item_w_cm + 0.20
                # Blank ends 5mm before next column
                blank_end_abs = next_c - 0.50
                blank_w = max(0.60, blank_end_abs - blank_start_abs)
                num_u = max(3, int(round(blank_w / u_width_cm)))

                if num_prefix:
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
                    sel.TypeText(f"{num_prefix} ")
                    sel.Font.Bold = 0
                    sel.Font.Color = 0

                from uln_parser import parse_inline_spans as _pis
                body_spans = _pis(word_part)
                self.write_inline_spans(sel, body_spans)

                # Blank starts exactly 2mm after text (rendered as a space + calculated underscores)
                sel.Font.Bold = 0
                sel.Font.Italic = 0
                sel.Font.Underline = 0
                sel.Font.Color = 0
                sel.TypeText(" ")
                sel.TypeText("_" * num_u)
            else:
                if num_prefix:
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
                    sel.TypeText(f"{num_prefix} ")
                    sel.Font.Bold = 0
                    sel.Font.Color = 0

                    from uln_parser import parse_inline_spans as _pis
                    body_spans = _pis(body_content)
                    self.write_inline_spans(sel, body_spans)
                else:
                    c_spans = block.cols_spans[c_idx] if (block.cols_spans and c_idx < len(block.cols_spans)) else []
                    if c_spans:
                        self.write_inline_spans(sel, c_spans)
                    else:
                        from uln_parser import parse_inline_spans as _pis
                        self.write_inline_spans(sel, _pis(col_raw))

        sel.TypeParagraph()
        self.last_rendered_tag = "TAB3" if num_cols == 3 else f"TAB{num_cols}"
