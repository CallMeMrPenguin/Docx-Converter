import os
import re
from typing import List, Optional
from uln_parser import ULNBlock, parse_inline_spans, PicInfo, parse_pic_tag
from renderer_utils import extract_question_prefix_and_body
from renderers.common.units_and_colors import cm_to_pt, pt_to_cm, parse_color_to_rgb_int


class TabColumnsRendererMixin:
    """Renders multi-column paragraph tab stops: [TAB2], [TAB3], [TAB4], and [TAB]."""

    def _insert_sign_picture_shape(self, sel, doc, pic_info: PicInfo, left_pt: float, width_pt: float, height_pt: float):
        """Helper to insert picture shape or placeholder securely anchored to current paragraph."""
        get_next_path = getattr(self, "get_next_image_path", None)
        target_path = get_next_path(pic_info) if get_next_path else None
        font_name = getattr(self, "font_name", "Times New Roman")

        if target_path and os.path.exists(target_path):
            try:
                norm_path = os.path.normpath(os.path.abspath(target_path))
                inline_shp = sel.InlineShapes.AddPicture(norm_path)
                shp = inline_shp.ConvertToShape()
                shp.RelativeHorizontalPosition = 0  # wdRelativeHorizontalPositionMargin = 0
                shp.RelativeVerticalPosition = 2    # wdRelativeVerticalPositionParagraph = 2
                shp.Left = left_pt
                shp.Top = 0.0
                shp.Width = width_pt
                shp.Height = height_pt
                shp.WrapFormat.Type = 3  # wdWrapNone = 3 (In front of text, 0 collision)
                shp.LockAnchor = True
                return
            except Exception as e:
                print(f"[ULNRenderer] Warning adding sign picture via ConvertToShape: {e}")

        # Fallback Placeholder Box
        try:
            curr_para_rng = sel.Paragraphs(1).Range
            shp = doc.Shapes.AddShape(
                5,  # msoShapeRoundedRectangle = 5
                left_pt,
                0.0,
                width_pt,
                height_pt,
                curr_para_rng
            )
            shp.RelativeHorizontalPosition = 0
            shp.RelativeVerticalPosition = 2
            shp.Left = left_pt
            shp.Top = 0.0
            shp.Width = width_pt
            shp.Height = height_pt
            shp.WrapFormat.Type = 3  # wdWrapNone = 3
            shp.Fill.ForeColor.RGB = 16316664  # Soft light grey
            shp.Line.ForeColor.RGB = 8421504   # Medium grey border
            shp.Line.Weight = 0.75
            shp.LockAnchor = True
            tf = shp.TextFrame
            tf.MarginLeft = 2.0
            tf.MarginRight = 2.0
            tf.MarginTop = 2.0
            tf.MarginBottom = 2.0
            tf.TextRange.Font.Name = font_name
            tf.TextRange.Font.Size = 9.0
            tf.TextRange.Font.Bold = 1
            tf.TextRange.Font.Color = 5263440  # Dark slate
            tf.TextRange.ParagraphFormat.Alignment = 1
            tf.TextRange.Text = f"[ 🖼️ {pic_info.description or 'SIGN IMAGE'} ]"
        except Exception as e:
            print(f"[ULNRenderer] Warning adding placeholder shape: {e}")

    def compute_group_sign_mcq_params(self, blocks: List[ULNBlock], printable_width_cm: float):
        """
        Pre-scans all Sign MCQ questions in a container group to calculate uniform picture dimensions
        (width and height) across the entire exercise section.
        - Measures option text lengths using GDI to detect any wrapped option lines.
        - Scales all pictures uniformly in proportion to the question with the maximum number of lines (including wrapped lines).
        - Enforces a minimum 5.0 mm gap, preferring 10.0 mm gap between text options and the picture on the right.
        """
        font_name = getattr(self, "font_name", "Times New Roman")
        font_size = getattr(self, "font_size", 12.0)
        from renderers.common.typography import get_gdi_text_measurer
        measurer = get_gdi_text_measurer()

        sign_questions = []
        i = 0
        while i < len(blocks):
            b = blocks[i]
            next_b = blocks[i + 1] if i + 1 < len(blocks) else None

            is_sign_mcq = False
            q_raw = ""
            if next_b and next_b.tag == "OPT":
                if b.content and re.search(r'\[PIC(?::.*?)?\]', b.content, re.IGNORECASE):
                    is_sign_mcq = True
                    q_raw = re.sub(r'\s*\[PIC(?::.*?)?\]\s*$', '', b.content, flags=re.IGNORECASE).strip()
                elif b.tag.startswith("TAB") and (b.col1 or b.col2):
                    c1 = b.col1 or ""
                    c2 = b.col2 or ""
                    if re.search(r'\[PIC(?::.*?)?\]', c1, re.IGNORECASE) or re.search(r'\[PIC(?::.*?)?\]', c2, re.IGNORECASE):
                        is_sign_mcq = True
                        q_raw = c2 if re.search(r'\[PIC(?::.*?)?\]', c1, re.IGNORECASE) else c1

            if is_sign_mcq:
                sign_questions.append((q_raw, next_b))
                i += 2
            else:
                i += 1

        if not sign_questions:
            return None, None, None

        est_pic_w_cm = 2.40
        preferred_gap_cm = 1.00  # 10 mm preferred gap
        avail_text_w_cm = max(6.0, printable_width_cm - est_pic_w_cm - preferred_gap_cm - 0.70)
        avail_text_w_pt = cm_to_pt(avail_text_w_cm)

        max_visual_lines = 3
        for q_raw, opt_b in sign_questions:
            pref, delim, q_num, c_body = extract_question_prefix_and_body(q_raw)
            has_body = bool(c_body.strip())
            q_lines = 1 if has_body else 0

            raw_opt_text = opt_b.content.strip()
            if '|' in raw_opt_text:
                opt_items = [x.strip() for x in raw_opt_text.split('|') if x.strip()]
            elif '\n' in raw_opt_text:
                opt_items = [x.strip() for x in raw_opt_text.split('\n') if x.strip()]
            else:
                from renderer_utils import split_line_into_option_items
                opt_items = split_line_into_option_items(raw_opt_text)

            total_opt_lines = 0
            for idx_o, item in enumerate(opt_items):
                clean_item = self.strip_markup_for_measurement(item)
                let_str = chr(65 + idx_o) if idx_o < 26 else str(idx_o + 1)
                full_str = f"{let_str}. {clean_item}"
                w_pt = measurer.measure_text_pt(full_str, font_name=font_name, font_size_pt=font_size, is_bold=False) * 1.15
                lines_for_this_opt = max(1, int(w_pt // avail_text_w_pt) + (1 if w_pt % avail_text_w_pt > 0 else 0))
                total_opt_lines += lines_for_this_opt

            total_lines_for_q = q_lines + total_opt_lines
            if total_lines_for_q > max_visual_lines:
                max_visual_lines = total_lines_for_q

        # Calculate uniform picture size matching the max visual lines
        uniform_h_cm = max(2.20, min(5.50, max_visual_lines * 0.72))
        uniform_w_cm = min(3.20, uniform_h_cm)

        # Gap calculation: maintain minimum 5.0 mm, preferring 10.0 mm (1.0 cm)
        rem_gap_space = printable_width_cm - uniform_w_cm - 8.0
        gap_cm = max(0.50, min(1.00, rem_gap_space)) if rem_gap_space < 1.00 else 1.00

        return uniform_w_cm, uniform_h_cm, gap_cm

    def render_side_by_side_pic_mcq(self, sel, doc, word, tab_block: ULNBlock, opt_block: ULNBlock, printable_width_cm: float):
        """
        Renders a 2-column Picture MCQ question using 100% PURE native paragraph tab stops and indents (NO TABLES).
        - Picture is ALWAYS positioned on the RIGHT side, anchored via ConvertToShape so it NEVER bunches up or drifts.
        - Uses synchronized uniform picture dimensions (width and height) across the entire exercise section.
        - Enforces minimum 5.0 mm, preferred 10.0 mm gap between text and picture.
        - Option text naturally wraps if long within RightIndent.
        - Word native Numbered List is applied with hanging indent, keeping option letters A., B., C., D. 100% vertically aligned.
        """
        font_name = getattr(self, "font_name", "Times New Roman")
        font_size = getattr(self, "font_size", 12.0)
        question_color = getattr(self, "question_color", None)
        opt_color = getattr(self, "opt_color", None)
        printable_width_pt = cm_to_pt(printable_width_cm)

        c1_str = tab_block.col1.strip()
        c2_str = tab_block.col2.strip()

        pic_in_c1 = bool(re.search(r'\[PIC(?::.*?)?\]', c1_str, re.IGNORECASE))
        pic_in_c2 = bool(re.search(r'\[PIC(?::.*?)?\]', c2_str, re.IGNORECASE))

        if pic_in_c1:
            pic_str = c1_str
            q_raw = c2_str
        else:
            pic_str = c2_str
            q_raw = c1_str

        pic_info = parse_pic_tag(pic_str) or PicInfo(description="Sign / Picture", pos="right", size="medium")

        # 1. Parse Options
        raw_opt_text = opt_block.content.strip()
        if '|' in raw_opt_text:
            opt_items = [x.strip() for x in raw_opt_text.split('|') if x.strip()]
        elif '\n' in raw_opt_text:
            opt_items = [x.strip() for x in raw_opt_text.split('\n') if x.strip()]
        else:
            from renderer_utils import split_line_into_option_items
            opt_items = split_line_into_option_items(raw_opt_text)

        normalized_opts = []
        for idx_opt, item in enumerate(opt_items):
            m_let = re.match(r'^\s*(?:(?:\*\*|\*|\[|\(?)*([a-zA-Z][\.\)])(?:\*\*|\*|\]|\}|\{u\}|\))*)\s*(.*)$', item)
            if m_let:
                let_part = m_let.group(1).rstrip('.)')
                body_part = m_let.group(2).strip()
                normalized_opts.append((let_part, body_part))
            else:
                let_part = chr(65 + idx_opt) if idx_opt < 26 else str(idx_opt + 1)
                normalized_opts.append((let_part, item))

        # 2. Extract question number and body
        pref, delim, q_num, c_body = extract_question_prefix_and_body(q_raw)
        q_display_body = c_body.strip() if c_body else ""

        # 3. Dynamic Picture Dimensions (from pre-computed uniform group params or local calculation)
        grp_w = getattr(self, "current_group_sign_pic_w_cm", None)
        grp_h = getattr(self, "current_group_sign_pic_h_cm", None)
        grp_gap = getattr(self, "current_group_sign_pic_gap_cm", None)

        if grp_w is not None and grp_h is not None:
            pic_w_cm = grp_w
            pic_h_cm = grp_h
            gap_cm = grp_gap if grp_gap is not None else 1.00
        else:
            num_opts = len(normalized_opts)
            total_lines = num_opts + (1 if q_display_body else 0)
            pic_h_cm = max(2.2, min(5.0, total_lines * 0.72))
            pic_w_cm = min(3.2, pic_h_cm)
            gap_cm = 1.00

        pic_w_pt = cm_to_pt(pic_w_cm)
        pic_h_pt = cm_to_pt(pic_h_cm)
        gap_pt = cm_to_pt(gap_cm)
        right_margin_indent = pic_w_pt + gap_pt
        pic_left_pt = printable_width_pt - pic_w_pt

        opt_color_int = parse_color_to_rgb_int(opt_color)

        # Standard hanging indent so options A, B, C are 100% vertically aligned
        indent_pt = cm_to_pt(0.60)

        # Calculate minimum vertical height needed for the question block to fit the picture
        total_items_count = len(normalized_opts) + (1 if q_display_body else 0)
        approx_text_block_h_pt = total_items_count * (font_size * 1.35) + 10.0
        extra_space_after_pt = max(2.0, pic_h_pt - approx_text_block_h_pt + 6.0)

        # ── Render Text on LEFT with Picture on RIGHT ───────────────
        if q_display_body:
            # Line 1: Question sentence
            sel.ParagraphFormat.LeftIndent = indent_pt
            sel.ParagraphFormat.FirstLineIndent = -indent_pt
            sel.ParagraphFormat.RightIndent = right_margin_indent
            sel.ParagraphFormat.SpaceBefore = 8
            sel.ParagraphFormat.SpaceAfter = 3
            sel.ParagraphFormat.KeepWithNext = True

            if q_num is not None:
                num_fmt = self.get_effective_number_format(pref, delim)
                self.apply_native_numbered_list(word, sel, q_num=q_num, number_format=num_fmt)
                sel.ParagraphFormat.LeftIndent = indent_pt
                sel.ParagraphFormat.FirstLineIndent = -indent_pt
                sel.ParagraphFormat.RightIndent = right_margin_indent

            self.write_inline_spans(sel, parse_inline_spans(q_display_body))
            self._insert_sign_picture_shape(sel, doc, pic_info, pic_left_pt, pic_w_pt, pic_h_pt)
            sel.TypeParagraph()

            try:
                sel.Range.ListFormat.RemoveNumbers()
            except Exception:
                pass

            for idx_opt, (let, body) in enumerate(normalized_opts):
                is_last_opt = (idx_opt == len(normalized_opts) - 1)
                sel.ParagraphFormat.LeftIndent = indent_pt + cm_to_pt(0.35)
                sel.ParagraphFormat.RightIndent = right_margin_indent
                sel.ParagraphFormat.FirstLineIndent = 0
                sel.ParagraphFormat.SpaceBefore = 2
                sel.ParagraphFormat.SpaceAfter = extra_space_after_pt if is_last_opt else 2
                sel.ParagraphFormat.KeepWithNext = True

                sel.Font.Name = font_name
                sel.Font.Size = font_size
                sel.Font.Bold = 1
                sel.Font.Italic = 0
                sel.Font.Underline = 0
                sel.Font.Color = opt_color_int if opt_color_int is not None else 0
                sel.TypeText(f"{let}. ")
                sel.Font.Bold = 0
                sel.Font.Color = 0

                self.write_inline_spans(sel, parse_inline_spans(body))
                sel.TypeParagraph()

        else:
            # Options-only: Line 1 has Native Numbered List + Option A
            for idx_opt, (let, body) in enumerate(normalized_opts):
                is_last_opt = (idx_opt == len(normalized_opts) - 1)
                sel.ParagraphFormat.RightIndent = right_margin_indent
                sel.ParagraphFormat.SpaceBefore = 8 if idx_opt == 0 else 2
                sel.ParagraphFormat.SpaceAfter = extra_space_after_pt if is_last_opt else 2
                sel.ParagraphFormat.KeepWithNext = True

                if idx_opt == 0:
                    sel.ParagraphFormat.LeftIndent = indent_pt
                    sel.ParagraphFormat.FirstLineIndent = -indent_pt
                    if q_num is not None:
                        num_fmt = self.get_effective_number_format(pref, delim)
                        self.apply_native_numbered_list(word, sel, q_num=q_num, number_format=num_fmt)
                        sel.ParagraphFormat.LeftIndent = indent_pt
                        sel.ParagraphFormat.FirstLineIndent = -indent_pt
                        sel.ParagraphFormat.RightIndent = right_margin_indent
                else:
                    try:
                        sel.Range.ListFormat.RemoveNumbers()
                    except Exception:
                        pass
                    sel.ParagraphFormat.LeftIndent = indent_pt
                    sel.ParagraphFormat.FirstLineIndent = 0

                sel.Font.Name = font_name
                sel.Font.Size = font_size
                sel.Font.Bold = 1
                sel.Font.Italic = 0
                sel.Font.Underline = 0
                sel.Font.Color = opt_color_int if opt_color_int is not None else 0
                sel.TypeText(f"{let}. ")
                sel.Font.Bold = 0
                sel.Font.Color = 0

                self.write_inline_spans(sel, parse_inline_spans(body))

                if idx_opt == 0:
                    self._insert_sign_picture_shape(sel, doc, pic_info, pic_left_pt, pic_w_pt, pic_h_pt)

                sel.TypeParagraph()

        sel.ParagraphFormat.LeftIndent = 0
        sel.ParagraphFormat.RightIndent = 0
        sel.ParagraphFormat.FirstLineIndent = 0
        try:
            sel.Range.ListFormat.RemoveNumbers()
        except Exception:
            pass
        self.last_rendered_tag = "OPT"

    def render_tab2_group(self, sel, doc, word, tab2_group: List[ULNBlock], printable_width_cm: float):
        """
        Renders 2-column items [TAB2] using 100% pure native Word Paragraph Tab Stops (NO TABLE).
        Handles:
        1. Blank lines (<blank> / ______): Right-aligned Tab Stop (Alignment=2) with RightIndent protection (5mm gap).
        2. Matching questions (Col 1 | Col 2): Left-aligned Tab Stop (Alignment=0) calculated from max width of Column 1.
        3. Headers (Col A | Col B): Center-aligned Tab Stops (Alignment=1).
        4. Pictures in Column 1 or Column 2: Aligned via Tab Stops.
        """
        if not tab2_group:
            return

        font_name = getattr(self, "font_name", "Times New Roman")
        font_size = getattr(self, "font_size", 12.0)
        question_color = getattr(self, "question_color", None)
        opt_color = getattr(self, "opt_color", None)

        group_first_c1 = tab2_group[0].col1 if tab2_group else ""
        base_indent_cm = 0.5 if "P1" in group_first_c1 else (1.0 if "P2" in group_first_c1 else 0.0)

        # 1. Determine if this TAB2 group is Question + Answer Blank
        has_any_blank = any(
            bool(re.search(r'^\s*(?:Answer:\s*)?(?:_{2,}|<blank>|\[BLANK\])\s*$', b.col2, re.IGNORECASE))
            for b in tab2_group
        )

        # 2. Calculate Tab Stop Position for matching/text layout
        if not has_any_blank:
            c1_clean_texts = []
            for b in tab2_group:
                raw_t = re.sub(r'^\s*\[(?:P0|P1|P2|INS)\]\s*', '', b.col1, flags=re.IGNORECASE).replace('#', '').strip()
                clean_t = self.strip_markup_for_measurement(raw_t)
                c1_clean_texts.append(clean_t)
            max_c1_w_pt = max((self.measure_text_width_pt(doc, t, font_name, font_size, is_bold=False) for t in c1_clean_texts), default=100.0)
            c1_word_w_cm = pt_to_cm(max_c1_w_pt) * 1.10

            c2_clean_texts = []
            for b in tab2_group:
                clean_c2 = self.strip_markup_for_measurement(b.col2.strip())
                c2_clean_texts.append(clean_c2)
            max_c2_w_pt = max((self.measure_text_width_pt(doc, t, font_name, font_size, is_bold=False) for t in c2_clean_texts), default=50.0)
            c2_word_w_cm = pt_to_cm(max_c2_w_pt) * 1.10

            min_gap_cm = 0.50
            tab_min_cm = base_indent_cm + c1_word_w_cm + min_gap_cm
            tab_max_cm = printable_width_cm - c2_word_w_cm - 0.10

            if tab_max_cm >= tab_min_cm:
                col2_tab_pos_cm = max(tab_min_cm, (base_indent_cm + printable_width_cm) / 2.0)
                if col2_tab_pos_cm > tab_max_cm:
                    col2_tab_pos_cm = tab_max_cm
            else:
                col2_tab_pos_cm = max(base_indent_cm + 4.0, tab_min_cm)

            if col2_tab_pos_cm >= printable_width_cm - 1.5:
                col2_tab_pos_cm = printable_width_cm - 2.5
        else:
            col2_tab_pos_cm = printable_width_cm

        # 3. Render each item using pure native Word Paragraph Tab Stops
        for idx, block in enumerate(tab2_group):
            last_tag = getattr(self, "last_rendered_tag", None)
            space_before_tab2 = 14 if (last_tag == "BOX") else 3
            sel.ParagraphFormat.SpaceBefore = space_before_tab2
            sel.ParagraphFormat.SpaceAfter = 3
            sel.ParagraphFormat.KeepWithNext = False
            sel.ParagraphFormat.PageBreakBefore = False
            sel.ParagraphFormat.LeftIndent = cm_to_pt(base_indent_cm)
            sel.ParagraphFormat.RightIndent = 0
            sel.ParagraphFormat.FirstLineIndent = 0

            # Detect header row in TAB2 (e.g. A | B or Column A | Column B)
            is_header_row = (idx == 0 and len(block.col1.strip()) <= 15 and len(block.col2.strip()) <= 15 and not re.search(r'\d', block.col1) and not has_any_blank)

            col2_is_blank = bool(re.match(r'^\s*(?:Answer:\s*)?(?:_{2,}|<blank>|\[BLANK\])\s*$', block.col2, re.IGNORECASE))

            if is_header_row:
                col1_needed_cm = col2_tab_pos_cm - base_indent_cm
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
                # ── Right-aligned Tab Stop (Alignment=2) with Protected RightIndent ──
                # Blank width = 2.8 cm, Min gap = 0.50 cm (5.0 mm).
                # RightIndent = 2.8cm + 0.50cm = 3.30cm restricts all text in Column 1 to wrap
                # strictly before (printable_width_cm - 3.30cm), guaranteeing the text NEVER
                # encroaches on Column 2 and always maintains at least a 5.0 mm physical gap!
                blank_w_cm = 2.8
                min_gap_cm = 0.50  # 5.0 mm strictly preserved
                right_indent_cm = blank_w_cm + min_gap_cm

                sel.ParagraphFormat.RightIndent = cm_to_pt(right_indent_cm)
                sel.ParagraphFormat.TabStops.ClearAll()
                sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(printable_width_cm), Alignment=2)

                pref, delim, q_num, c1_body = extract_question_prefix_and_body(block.col1)
                if q_num is not None:
                    pref_str = pref if pref else ""
                    delim_char = delim if delim else "."
                    num_prefix_str = f"{pref_str}{q_num}{delim_char} "
                    sel.Font.Name = font_name
                    sel.Font.Size = font_size
                    sel.Font.Bold = 1
                    sel.Font.Italic = 0
                    sel.Font.Underline = 0
                    q_color_int = parse_color_to_rgb_int(question_color)
                    sel.Font.Color = q_color_int if q_color_int is not None else 0
                    sel.TypeText(num_prefix_str)
                    sel.Font.Bold = 0
                    sel.Font.Color = 0

                    c1_spans = parse_inline_spans(c1_body.strip())
                    self.write_inline_spans(sel, c1_spans)
                else:
                    self.write_inline_spans(sel, block.col1_spans)

                sel.Font.Color = 0
                sel.Font.Bold = 0
                sel.Font.Underline = 0
                sel.TypeText("\t___________")
                sel.TypeParagraph()
                sel.ParagraphFormat.RightIndent = 0

            else:
                # ── Standard 2-Column Matching Layout ──
                sel.ParagraphFormat.TabStops.ClearAll()
                sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(col2_tab_pos_cm), Alignment=0)

                pref, delim, q_num, c1_body = extract_question_prefix_and_body(block.col1)
                if q_num is not None:
                    pref_str = pref if pref else ""
                    delim_char = delim if delim else "."
                    num_prefix_str = f"{pref_str}{q_num}{delim_char} "
                    sel.Font.Name = font_name
                    sel.Font.Size = font_size
                    sel.Font.Bold = 1
                    sel.Font.Italic = 0
                    sel.Font.Underline = 0
                    q_color_int = parse_color_to_rgb_int(question_color)
                    sel.Font.Color = q_color_int if q_color_int is not None else 0
                    sel.TypeText(num_prefix_str)
                    sel.Font.Bold = 0
                    sel.Font.Color = 0

                    c1_spans = parse_inline_spans(c1_body.strip())
                    self.write_inline_spans(sel, c1_spans)
                else:
                    self.write_inline_spans(sel, block.col1_spans)

                sel.TypeText("\t")

                col2_trim = block.col2.strip()
                m_opt = re.match(r'^\s*(?:(?:\*\*|\*|\[|\(?)*([a-zA-Z])[\.\)](?:\*\*|\*|\]|\}|\{u\}|\))*)\s+(.*)$', col2_trim)
                pref2, delim2, q_num2, c2_body = extract_question_prefix_and_body(block.col2)

                if q_num2 is not None:
                    pref_str2 = pref2 if pref2 else ""
                    delim_char2 = delim2 if delim2 else "."
                    num_prefix_str2 = f"{pref_str2}{q_num2}{delim_char2} "
                    sel.Font.Name = font_name
                    sel.Font.Size = font_size
                    sel.Font.Bold = 1
                    sel.Font.Italic = 0
                    sel.Font.Underline = 0
                    q_color_int = parse_color_to_rgb_int(question_color)
                    sel.Font.Color = q_color_int if q_color_int is not None else 0
                    sel.TypeText(num_prefix_str2)
                    sel.Font.Bold = 0
                    sel.Font.Color = 0

                    c2_spans = parse_inline_spans(c2_body.strip())
                    self.write_inline_spans(sel, c2_spans)
                elif m_opt and not block.pic:
                    opt_let = f"{m_opt.group(1).upper()}."
                    opt_body = m_opt.group(2).strip()

                    sel.Font.Name = font_name
                    sel.Font.Size = font_size
                    sel.Font.Bold = 1
                    sel.Font.Italic = 0
                    sel.Font.Underline = 0
                    opt_color_int = parse_color_to_rgb_int(opt_color)
                    sel.Font.Color = opt_color_int if opt_color_int is not None else 0
                    sel.TypeText(f"{opt_let} ")
                    sel.Font.Bold = 0
                    sel.Font.Color = 0

                    self.write_inline_spans(sel, parse_inline_spans(opt_body))
                else:
                    self.write_inline_spans(sel, block.col2_spans)

                sel.TypeParagraph()

        sel.ParagraphFormat.LeftIndent = 0
        sel.ParagraphFormat.FirstLineIndent = 0
        sel.ParagraphFormat.TabStops.ClearAll()
        self.last_rendered_tag = "TAB2"

    def render_tab_multi(self, sel, doc, block: ULNBlock, idx_block: int, blocks: List[ULNBlock], printable_width_cm: float):
        """
        Renders 3-column (TAB3), 4-column (TAB4), or N-column (TAB) side-by-side items
        using native Word paragraph tab stops with optimal gap calculations and aligned blanks.
        """
        num_cols = len(block.cols) if block.cols else (3 if block.tag == "TAB3" else (4 if block.tag == "TAB4" else 3))
        font_name = getattr(self, "font_name", "Times New Roman")
        font_size = getattr(self, "font_size", 12.0)
        question_color = getattr(self, "question_color", None)

        # 1. Collect full group of consecutive TAB blocks with the same column count
        group_start = idx_block
        while group_start > 0 and blocks[group_start - 1].tag.startswith("TAB") and blocks[group_start - 1].tag != "TABLE" and (len(blocks[group_start - 1].cols) if blocks[group_start - 1].cols else 0) == num_cols:
            group_start -= 1

        tab_group = []
        lookahead = group_start
        while lookahead < len(blocks) and blocks[lookahead].tag.startswith("TAB") and blocks[lookahead].tag != "TABLE" and (len(blocks[lookahead].cols) if blocks[lookahead].cols else 0) == num_cols:
            tab_group.append(blocks[lookahead])
            lookahead += 1

        group_first_c1 = tab_group[0].cols[0] if (tab_group and tab_group[0].cols) else (block.cols[0] if block.cols else "")
        base_indent_cm = 0.5 if "P1" in group_first_c1 else (1.0 if "P2" in group_first_c1 else 0.0)

        has_blanks = any(
            re.search(r'[_]{2,}|<(?:blank|BLANK)>|\[(?:blank|BLANK)\]', c)
            for b in tab_group for c in (b.cols or [])
        )

        avail_w_cm = printable_width_cm - base_indent_cm
        even_step_cm = avail_w_cm / float(num_cols)

        col_starts_cm = [base_indent_cm + (i * even_step_cm) for i in range(num_cols)]

        u_pt = self.measure_text_width_pt(doc, '_', font_name, font_size, is_bold=False)
        u_width_cm = max(0.20, pt_to_cm(u_pt) * 1.15)

        tab_stops_cm = [col_starts_cm[i] for i in range(1, num_cols)]

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

        for c_idx, col_raw in enumerate(block.cols):
            if c_idx > 0:
                sel.TypeText("\t")

            clean_col = col_raw.strip()
            m_num = re.match(r'^\s*#?(\d+[\.\)])\s*(.*)$', clean_col)
            num_prefix = m_num.group(1) if m_num else ""
            body_content = m_num.group(2) if m_num else clean_col

            if has_blanks and re.search(r'[_]{2,}|<(?:blank|BLANK)>|\[(?:blank|BLANK)\]', body_content):
                word_part = re.sub(r'[_]{2,}|<(?:blank|BLANK)>|\[(?:blank|BLANK)\]', '', body_content).strip()

                num_str = f"{num_prefix} " if num_prefix else ""
                full_item_str = f"{num_str}{word_part}".strip()
                w_pt = self.measure_text_width_pt(doc, full_item_str, font_name, font_size, is_bold=False)
                item_w_cm = pt_to_cm(w_pt) * 1.15

                start_c = col_starts_cm[c_idx]
                next_c = col_starts_cm[c_idx + 1] if (c_idx + 1 < num_cols) else printable_width_cm

                blank_start_abs = start_c + item_w_cm + 0.20
                blank_end_abs = next_c - 0.50
                blank_w = max(0.60, blank_end_abs - blank_start_abs)
                num_u = max(3, int(round(blank_w / u_width_cm)))

                if num_prefix:
                    sel.Font.Name = font_name
                    sel.Font.Size = font_size
                    sel.Font.Bold = 1
                    sel.Font.Italic = 0
                    sel.Font.Underline = 0
                    q_color_int = parse_color_to_rgb_int(question_color)
                    sel.Font.Color = q_color_int if q_color_int is not None else 0
                    sel.TypeText(f"{num_prefix} ")
                    sel.Font.Bold = 0
                    sel.Font.Color = 0

                body_spans = parse_inline_spans(word_part)
                self.write_inline_spans(sel, body_spans)

                sel.Font.Bold = 0
                sel.Font.Italic = 0
                sel.Font.Underline = 0
                sel.Font.Color = 0
                sel.TypeText(" ")
                sel.TypeText("_" * num_u)
            else:
                if num_prefix:
                    sel.Font.Name = font_name
                    sel.Font.Size = font_size
                    sel.Font.Bold = 1
                    sel.Font.Italic = 0
                    sel.Font.Underline = 0
                    q_color_int = parse_color_to_rgb_int(question_color)
                    sel.Font.Color = q_color_int if q_color_int is not None else 0
                    sel.TypeText(f"{num_prefix} ")
                    sel.Font.Bold = 0
                    sel.Font.Color = 0

                    body_spans = parse_inline_spans(body_content)
                    self.write_inline_spans(sel, body_spans)
                else:
                    c_spans = block.cols_spans[c_idx] if (block.cols_spans and c_idx < len(block.cols_spans)) else []
                    if c_spans:
                        self.write_inline_spans(sel, c_spans)
                    else:
                        self.write_inline_spans(sel, parse_inline_spans(col_raw))

        sel.TypeParagraph()
        self.last_rendered_tag = "TAB3" if num_cols == 3 else f"TAB{num_cols}"
