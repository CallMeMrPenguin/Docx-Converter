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

        # Calculate uniform picture size matching standard 17.5x25mm sign dimensions
        uniform_w_cm = 2.50
        uniform_h_cm = 1.75
        gap_cm = 0.50

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
            gap_cm = grp_gap if grp_gap is not None else 0.50
        else:
            pic_w_cm = 2.50
            pic_h_cm = 1.75
            gap_cm = 0.50

        pic_w_pt = cm_to_pt(pic_w_cm)
        pic_h_pt = cm_to_pt(pic_h_cm)
        gap_pt = cm_to_pt(gap_cm)
        right_margin_indent = pic_w_pt + gap_pt
        pic_left_pt = printable_width_pt - pic_w_pt

        opt_color_int = parse_color_to_rgb_int(opt_color)

        # Standard hanging indent so options A, B, C are 100% vertically aligned (5mm)
        indent_pt = cm_to_pt(0.50)

        # Calculate minimum vertical height needed for the question block to fit the picture
        total_items_count = len(normalized_opts) + (1 if q_display_body else 0)
        approx_text_block_h_pt = total_items_count * (font_size * 1.35) + 4.0
        extra_space_after_pt = max(4.0, pic_h_pt - approx_text_block_h_pt + 8.0)

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
                sel.ParagraphFormat.KeepWithNext = not is_last_opt

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
                sel.ParagraphFormat.KeepWithNext = not is_last_opt

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
        2. Matching questions (Col 1 | Col 2): Left-aligned Tab Stop (Alignment=0) with strict column-bound wrapping.
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

        # 1. Determine if this TAB2 group contains Pictures
        has_pic_in_c1 = any(
            bool(re.search(r'\[PIC(?::.*?)?\]', b.col1, re.IGNORECASE) or parse_pic_tag(b.col1) is not None)
            for b in tab2_group
        )
        has_pic_in_c2 = any(
            bool(re.search(r'\[PIC(?::.*?)?\]', b.col2, re.IGNORECASE) or parse_pic_tag(b.col2) is not None)
            for b in tab2_group
        )

        if has_pic_in_c1:
            # ── 2-Column Layout: Picture on Left (Col 1), Text on Right (Col 2) ──
            pic_w_cm = 2.50  # 25 mm width
            pic_h_cm = 1.75  # 17.5 mm height (17.5x25mm)
            min_gap_cm = 0.50  # Strict 5 mm gap between image and text

            # Measure question number prefix widths
            num_widths = []
            for b in tab2_group:
                pref, delim, q_num, _ = extract_question_prefix_and_body(b.col1)
                if q_num is not None:
                    pref_str = pref if pref else ""
                    delim_char = delim if delim else "."
                    num_str = f"{pref_str}{q_num}{delim_char} "
                    w_pt = self.measure_text_width_pt(doc, num_str, font_name, font_size, is_bold=True)
                    num_widths.append(pt_to_cm(w_pt) * 1.10)
                else:
                    num_widths.append(0.0)

            max_num_w_cm = max(num_widths, default=0.70)
            col1_total_w_cm = base_indent_cm + max_num_w_cm + pic_w_cm + 0.15
            col2_pos_cm = col1_total_w_cm + min_gap_cm

            for idx, block in enumerate(tab2_group):
                last_tag = getattr(self, "last_rendered_tag", None)
                sel.ParagraphFormat.SpaceBefore = 14 if (last_tag == "BOX" and idx == 0) else 3
                sel.ParagraphFormat.SpaceAfter = 3
                sel.ParagraphFormat.KeepWithNext = False
                sel.ParagraphFormat.PageBreakBefore = False

                # Hanging indent: guarantees multi-line wrapped text in Col 2 NEVER jumps into Col 1
                sel.ParagraphFormat.LeftIndent = cm_to_pt(col2_pos_cm)
                sel.ParagraphFormat.FirstLineIndent = -cm_to_pt(col2_pos_cm - base_indent_cm)
                sel.ParagraphFormat.RightIndent = 0
                sel.ParagraphFormat.TabStops.ClearAll()
                sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(col2_pos_cm), Alignment=0)

                # Line 1: Question Number Prefix
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

                # Line 1: Sign Picture (InlineShape, sized 17.5x25mm)
                pic_info = parse_pic_tag(block.col1) or parse_pic_tag(c1_body) or PicInfo(description="Sign", pos="center", size="small")
                self.current_tab2_pic_width_cm = pic_w_cm
                self.current_tab2_pic_height_cm = pic_h_cm
                self.render_pic(sel, doc, pic_info)
                self.current_tab2_pic_width_cm = None
                self.current_tab2_pic_height_cm = None

                # Advance to Column 2
                sel.TypeText("\t")

                # Column 2 Text (processes inline blanks e.g. <blank> -> ___________)
                c2_text = block.col2.strip()
                c2_spans = parse_inline_spans(c2_text)
                self.write_inline_spans(sel, c2_spans)
                sel.TypeParagraph()

            sel.ParagraphFormat.LeftIndent = 0
            sel.ParagraphFormat.FirstLineIndent = 0
            sel.ParagraphFormat.TabStops.ClearAll()
            self.last_rendered_tag = "TAB2"
            return

        elif has_pic_in_c2:
            # ── 2-Column Layout: Text on Left (Col 1), Picture on Right (Col 2) ──
            pic_w_cm = 2.50
            pic_h_cm = 1.75
            min_gap_cm = 0.50
            col_pic_pos_cm = printable_width_cm - pic_w_cm
            right_indent_cm = pic_w_cm + min_gap_cm

            for idx, block in enumerate(tab2_group):
                last_tag = getattr(self, "last_rendered_tag", None)
                sel.ParagraphFormat.SpaceBefore = 14 if (last_tag == "BOX" and idx == 0) else 3
                sel.ParagraphFormat.SpaceAfter = 3
                sel.ParagraphFormat.KeepWithNext = False
                sel.ParagraphFormat.PageBreakBefore = False

                sel.ParagraphFormat.LeftIndent = cm_to_pt(base_indent_cm)
                sel.ParagraphFormat.FirstLineIndent = 0
                sel.ParagraphFormat.RightIndent = cm_to_pt(right_indent_cm)
                sel.ParagraphFormat.TabStops.ClearAll()
                sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(col_pic_pos_cm), Alignment=0)

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

                pic_info = parse_pic_tag(block.col2) or PicInfo(description="Sign", pos="right", size="small")
                self.current_tab2_pic_width_cm = pic_w_cm
                self.current_tab2_pic_height_cm = pic_h_cm
                self.render_pic(sel, doc, pic_info)
                self.current_tab2_pic_width_cm = None
                self.current_tab2_pic_height_cm = None
                sel.TypeParagraph()

            sel.ParagraphFormat.LeftIndent = 0
            sel.ParagraphFormat.FirstLineIndent = 0
            sel.ParagraphFormat.RightIndent = 0
            sel.ParagraphFormat.TabStops.ClearAll()
            self.last_rendered_tag = "TAB2"
            return

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
            c1_word_w_cm = pt_to_cm(max_c1_w_pt) * 1.04

            c2_clean_texts = []
            for b in tab2_group:
                clean_c2 = self.strip_markup_for_measurement(b.col2.strip())
                c2_clean_texts.append(clean_c2)
            max_c2_w_pt = max((self.measure_text_width_pt(doc, t, font_name, font_size, is_bold=False) for t in c2_clean_texts), default=50.0)
            c2_word_w_cm = pt_to_cm(max_c2_w_pt) * 1.04

            min_gap_cm = 0.50  # Strict 5mm minimum gap
            safety_margin_cm = 0.35

            tab_min_cm = base_indent_cm + c1_word_w_cm + min_gap_cm
            tab_max_cm = printable_width_cm - c2_word_w_cm - safety_margin_cm

            if tab_max_cm >= tab_min_cm:
                # Case A: Both columns can fit on single line without wrapping.
                # Maximize distance between columns by allocating available slack without pushing Col 2 past tab_max_cm:
                slack = tab_max_cm - tab_min_cm
                col2_tab_pos_cm = tab_min_cm + (slack * 0.50)
            else:
                # Case B: Wrapping required. Evaluate layout strategies with line-budget simulation:
                def simulate_layout(pos_cm):
                    c1_avail_pt = max(cm_to_pt(1.5), cm_to_pt(pos_cm - base_indent_cm - min_gap_cm))
                    c2_avail_pt = max(cm_to_pt(1.5), cm_to_pt(printable_width_cm - pos_cm - safety_margin_cm))
                    tot_rows = 0
                    extra_c1 = 0
                    extra_c2 = 0
                    for b in tab2_group:
                        pref_s, delim_s, q_num_s, c1_b = extract_question_prefix_and_body(b.col1)
                        num_w = self.measure_text_width_pt(doc, f"{pref_s or ''}{q_num_s}{delim_s or '.'} ", font_name, font_size, is_bold=True) if q_num_s is not None else 0.0
                        l1 = len(self.wrap_text_into_lines(doc, (c1_b.strip() if q_num_s is not None else b.col1.strip()), c1_avail_pt - num_w))

                        m_opt = re.match(r'^\s*(?:(?:\*\*|\*|\[|\(?)*([a-zA-Z])[\.\)](?:\*\*|\*|\]|\}|\{u\}|\))*)\s+(.*)$', b.col2.strip())
                        opt_w = self.measure_text_width_pt(doc, f"{m_opt.group(1).upper()}. ", font_name, font_size, is_bold=True) if m_opt else 0.0
                        body2 = m_opt.group(2).strip() if m_opt else b.col2.strip()
                        l2 = len(self.wrap_text_into_lines(doc, body2, c2_avail_pt - opt_w))

                        tot_rows += max(l1, l2, 1)
                        extra_c1 += max(0, l1 - 1)
                        extra_c2 += max(0, l2 - 1)
                    return tot_rows, extra_c1, extra_c2

                # 1. Proportional Shared Wrapping Candidate
                total_w = c1_word_w_cm + c2_word_w_cm
                avail_for_both = printable_width_cm - base_indent_cm - min_gap_cm
                prop = c1_word_w_cm / total_w if total_w > 0 else 0.50
                clamped_prop = max(0.25, min(0.75, prop))
                pos_prop = base_indent_cm + (avail_for_both * clamped_prop) + min_gap_cm
                pos_prop = max(base_indent_cm + 2.5, min(printable_width_cm - 2.5, pos_prop))
                rows_prop, extra_c1_prop, extra_c2_prop = simulate_layout(pos_prop)

                # 2. Strategy 1: Prevent Col 1 wrap (Allocate full width to Col 1)
                pos_no_c1 = min(printable_width_cm - 3.0, max(base_indent_cm + 2.5, tab_min_cm))
                rows_no_c1, extra_c1_no_c1, extra_c2_no_c1 = simulate_layout(pos_no_c1)

                # 3. Strategy 2: Prevent Col 2 wrap (Allocate full width to Col 2)
                pos_no_c2 = max(base_indent_cm + 2.5, min(printable_width_cm - 3.0, tab_max_cm))
                rows_no_c2, extra_c1_no_c2, extra_c2_no_c2 = simulate_layout(pos_no_c2)

                # Evaluate decisions:
                col2_tab_pos_cm = pos_prop
                if extra_c1_prop < extra_c2_prop:
                    # Column 1 was wrapping fewer lines than Column 2.
                    # If unwrapping Column 1 does NOT create more total lines than proportional, unwrap Column 1 completely!
                    if rows_no_c1 <= rows_prop:
                        col2_tab_pos_cm = pos_no_c1
                elif extra_c2_prop < extra_c1_prop:
                    # Column 2 was wrapping fewer lines than Column 1.
                    # If unwrapping Column 2 does NOT create more total lines than proportional, unwrap Column 2 completely!
                    if rows_no_c2 <= rows_prop:
                        col2_tab_pos_cm = pos_no_c2

            col2_tab_pos_cm = max(base_indent_cm + 2.0, min(printable_width_cm - 2.0, col2_tab_pos_cm))
        else:
            col2_tab_pos_cm = printable_width_cm

        # 3. Render each item using pure native Word Paragraph Tab Stops
        for idx, block in enumerate(tab2_group):
            last_tag = getattr(self, "last_rendered_tag", None)
            space_before_tab2 = 14 if (last_tag == "BOX") else 3

            # Detect header row in TAB2 (e.g. A | B or Column A | Column B)
            is_header_row = (idx == 0 and len(block.col1.strip()) <= 15 and len(block.col2.strip()) <= 15 and not re.search(r'\d', block.col1) and not has_any_blank)

            col2_is_blank = bool(re.match(r'^\s*(?:Answer:\s*)?(?:_{2,}|<blank>|\[BLANK\])\s*$', block.col2, re.IGNORECASE))

            if is_header_row:
                col1_needed_cm = col2_tab_pos_cm - base_indent_cm
                col1_center_cm = base_indent_cm + (col1_needed_cm / 2.0)
                col2_center_cm = col2_tab_pos_cm + ((printable_width_cm - col2_tab_pos_cm) / 2.0)

                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.FirstLineIndent = 0
                sel.ParagraphFormat.SpaceBefore = space_before_tab2
                sel.ParagraphFormat.SpaceAfter = 3
                sel.ParagraphFormat.KeepWithNext = False
                sel.ParagraphFormat.PageBreakBefore = False
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
                blank_w_cm = 2.8
                min_gap_cm = 0.50
                right_indent_cm = blank_w_cm + min_gap_cm

                sel.ParagraphFormat.SpaceBefore = space_before_tab2
                sel.ParagraphFormat.SpaceAfter = 3
                sel.ParagraphFormat.KeepWithNext = False
                sel.ParagraphFormat.PageBreakBefore = False
                sel.ParagraphFormat.LeftIndent = cm_to_pt(base_indent_cm)
                sel.ParagraphFormat.FirstLineIndent = 0
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
                # ── Standard 2-Column Matching Layout with Strictly Contained Visual Wrapping ──
                col1_avail_pt = max(cm_to_pt(1.5), cm_to_pt(col2_tab_pos_cm - base_indent_cm - min_gap_cm))
                col2_avail_pt = max(cm_to_pt(1.5), cm_to_pt(printable_width_cm - col2_tab_pos_cm) - cm_to_pt(0.35))

                pref, delim, q_num, c1_body = extract_question_prefix_and_body(block.col1)
                num_prefix_str = ""
                num_w_pt = 0.0
                if q_num is not None:
                    pref_str = pref if pref else ""
                    delim_char = delim if delim else "."
                    num_prefix_str = f"{pref_str}{q_num}{delim_char} "
                    num_w_pt = self.measure_text_width_pt(doc, num_prefix_str, font_name, font_size, is_bold=True)
                    c1_body_avail_pt = max(cm_to_pt(1.5), col1_avail_pt - num_w_pt)
                    c1_lines = self.wrap_text_into_lines(doc, c1_body.strip(), c1_body_avail_pt)
                else:
                    c1_lines = self.wrap_text_into_lines(doc, block.col1.strip(), col1_avail_pt)

                col1_body_pos_cm = base_indent_cm + (pt_to_cm(num_w_pt) if num_prefix_str else 0.0)

                col2_trim = block.col2.strip()
                m_opt = re.match(r'^\s*(?:(?:\*\*|\*|\[|\(?)*([a-zA-Z])[\.\)](?:\*\*|\*|\]|\}|\{u\}|\))*)\s+(.*)$', col2_trim)
                pref2, delim2, q_num2, c2_body = extract_question_prefix_and_body(block.col2)

                num_prefix_str2 = ""
                opt_prefix_str = ""
                col2_body_pos_cm = col2_tab_pos_cm

                if q_num2 is not None:
                    pref_str2 = pref2 if pref2 else ""
                    delim_char2 = delim2 if delim2 else "."
                    num_prefix_str2 = f"{pref_str2}{q_num2}{delim_char2} "
                    num_w_pt2 = self.measure_text_width_pt(doc, num_prefix_str2, font_name, font_size, is_bold=True)
                    col2_body_pos_cm = col2_tab_pos_cm + pt_to_cm(num_w_pt2)
                    c2_body_avail_pt = max(cm_to_pt(1.5), col2_avail_pt - num_w_pt2)
                    c2_lines = self.wrap_text_into_lines(doc, c2_body.strip(), c2_body_avail_pt)
                elif m_opt and not block.pic:
                    opt_let = f"{m_opt.group(1).upper()}."
                    opt_prefix_str = f"{opt_let} "
                    opt_body = m_opt.group(2).strip()
                    opt_w_pt = self.measure_text_width_pt(doc, opt_prefix_str, font_name, font_size, is_bold=True)
                    col2_body_pos_cm = col2_tab_pos_cm + pt_to_cm(opt_w_pt)
                    c2_body_avail_pt = max(cm_to_pt(1.5), col2_avail_pt - opt_w_pt)
                    c2_lines = self.wrap_text_into_lines(doc, opt_body, c2_body_avail_pt)
                else:
                    c2_lines = self.wrap_text_into_lines(doc, col2_trim, col2_avail_pt)

                num_visual_rows = max(len(c1_lines), len(c2_lines), 1)

                for r in range(num_visual_rows):
                    is_last_row = (r == num_visual_rows - 1)
                    has_c1 = (r < len(c1_lines))
                    has_c2 = (r < len(c2_lines))

                    sel.ParagraphFormat.LeftIndent = cm_to_pt(base_indent_cm)
                    sel.ParagraphFormat.RightIndent = 0
                    sel.ParagraphFormat.FirstLineIndent = 0
                    sel.ParagraphFormat.TabStops.ClearAll()

                    if getattr(self, "is_inside_num_container", False):
                        sel.ParagraphFormat.SpaceBefore = 0
                        sel.ParagraphFormat.SpaceAfter = 0
                        sel.ParagraphFormat.LineSpacing = font_size * 1.16
                    else:
                        sel.ParagraphFormat.SpaceBefore = (space_before_tab2 if idx == 0 else 2) if r == 0 else 0
                        sel.ParagraphFormat.SpaceAfter = 3 if is_last_row else 0
                    sel.ParagraphFormat.KeepWithNext = not is_last_row

                    if r == 0:
                        # Row 0: Main Tab Stop for Column 2
                        sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(col2_tab_pos_cm), Alignment=0)

                        if num_prefix_str:
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

                        c1_spans = parse_inline_spans(c1_lines[0])
                        self.write_inline_spans(sel, c1_spans)

                        # Advance to Column 2 via Tab Stop
                        sel.TypeText("\t")

                        if opt_prefix_str:
                            sel.Font.Name = font_name
                            sel.Font.Size = font_size
                            sel.Font.Bold = 1
                            sel.Font.Italic = 0
                            sel.Font.Underline = 0
                            opt_color_int = parse_color_to_rgb_int(opt_color)
                            sel.Font.Color = opt_color_int if opt_color_int is not None else 0
                            sel.TypeText(opt_prefix_str)
                            sel.Font.Bold = 0
                            sel.Font.Color = 0

                            c2_spans = parse_inline_spans(c2_lines[0])
                            self.write_inline_spans(sel, c2_spans)
                        elif num_prefix_str2:
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

                            c2_spans = parse_inline_spans(c2_lines[0])
                            self.write_inline_spans(sel, c2_spans)
                        else:
                            c2_spans = parse_inline_spans(c2_lines[0])
                            self.write_inline_spans(sel, c2_spans)

                    else:
                        # Row 1+ (Wrapped lines aligned with pure native Word Tab Stops):
                        col2_dest_cm = col2_body_pos_cm if (opt_prefix_str or num_prefix_str2) else col2_tab_pos_cm

                        if has_c1:
                            if num_prefix_str and col1_body_pos_cm > base_indent_cm:
                                sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(col1_body_pos_cm), Alignment=0)
                            if has_c2:
                                sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(col2_dest_cm), Alignment=0)

                            if num_prefix_str and col1_body_pos_cm > base_indent_cm:
                                sel.TypeText("\t")

                            c1_spans = parse_inline_spans(c1_lines[r])
                            self.write_inline_spans(sel, c1_spans)

                            if has_c2:
                                sel.TypeText("\t")
                                c2_spans = parse_inline_spans(c2_lines[r])
                                self.write_inline_spans(sel, c2_spans)
                        else:
                            # Column 1 is empty on this wrapped line
                            if has_c2:
                                sel.ParagraphFormat.TabStops.Add(Position=cm_to_pt(col2_dest_cm), Alignment=0)
                                sel.TypeText("\t")
                                c2_spans = parse_inline_spans(c2_lines[r])
                                self.write_inline_spans(sel, c2_spans)

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
            pref_col, delim_col, q_num_col, body_content = extract_question_prefix_and_body(clean_col)
            num_prefix = f"{pref_col or ''}{q_num_col}{delim_col or '.'}" if q_num_col is not None else ""

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

            # Process inline blanks in body_content e.g. / <blank> / -> / ______ /
            if re.search(r'<(?:blank|BLANK)>|\[(?:blank|BLANK)\]|_{2,}', body_content):
                # If blank is embedded inside text or slashes (e.g. / <blank> /)
                if re.search(r'(?:\/|\w|\))\s*<(?:blank|BLANK)>|\[(?:blank|BLANK)\]|_{2,}', body_content) or body_content.rstrip().endswith('/'):
                    processed_body = re.sub(r'<(?:blank|BLANK)>|\[(?:blank|BLANK)\]|_{2,}', '______', body_content)
                    body_spans = parse_inline_spans(processed_body)
                    self.write_inline_spans(sel, body_spans)
                else:
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

                    body_spans = parse_inline_spans(word_part)
                    self.write_inline_spans(sel, body_spans)

                    sel.Font.Bold = 0
                    sel.Font.Italic = 0
                    sel.Font.Underline = 0
                    sel.Font.Color = 0
                    sel.TypeText(" ")
                    sel.TypeText("_" * num_u)
            else:
                body_spans = parse_inline_spans(body_content)
                self.write_inline_spans(sel, body_spans)

        sel.TypeParagraph()
        self.last_rendered_tag = "TAB3" if num_cols == 3 else f"TAB{num_cols}"
