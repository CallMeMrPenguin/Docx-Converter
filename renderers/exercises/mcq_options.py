import re
from typing import List, Tuple, Optional
from uln_parser import ULNBlock, parse_inline_spans
from renderer_utils import split_line_into_option_items, extract_question_prefix_and_body
from renderers.common.units_and_colors import cm_to_pt, parse_color_to_rgb_int
from renderers.common.typography import get_gdi_text_measurer


class McqOptionsRendererMixin:
    """Renders Multiple Choice Question options ([OPT]) and computes optimal column layouts."""

    def compute_group_option_params(self, opt_blocks: List[ULNBlock], printable_width_cm: float) -> Tuple[Optional[int], Optional[float]]:
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
            font_name = getattr(self, "font_name", "Times New Roman")
            font_size = getattr(self, "font_size", 12.0)
            max_w_pt = max(measurer.measure_text_pt(ci, font_name=font_name, font_size_pt=font_size, is_bold=False) for ci in clean_items) if clean_items else 0.0
            return cols, max_w_pt
        return None, None

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
        font_name = getattr(self, "font_name", "Times New Roman")
        font_size = getattr(self, "font_size", 12.0)
        item_widths_pt = [measurer.measure_text_pt(ci, font_name=font_name, font_size_pt=font_size, is_bold=False) * 1.15 for ci in clean_items]
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

        font_name = getattr(self, "font_name", "Times New Roman")
        font_size = getattr(self, "font_size", 12.0)

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
        local_max_w_pt = max(measurer.measure_text_pt(s, font_name=font_name, font_size_pt=font_size) for s in formatted_item_strings) if formatted_item_strings else 0.0

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

        if getattr(self, "is_inside_num_container", False):
            sel.ParagraphFormat.SpaceBefore = 0
            sel.ParagraphFormat.SpaceAfter = 0
            sel.ParagraphFormat.LineSpacing = getattr(self, "font_size", 12.0) * 1.16
        else:
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

                sel.Font.Name = font_name
                sel.Font.Size = font_size
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

                sel.Font.Name = font_name
                sel.Font.Size = font_size
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
