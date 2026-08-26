import re
import math
from typing import List, Tuple, Optional
from uln_parser import ULNBlock, parse_inline_spans
from renderers.common.units_and_colors import cm_to_pt


class BoxesRendererMixin:
    """Renders framed callout boxes, grammar formula boxes, and optimized Word Bank shapes ([BOX])."""

    def clean_box_item(self, text: str) -> str:
        """Strips inline ULN phonetic/formatting markup tags so text displays clean words in Word Bank."""
        if not text:
            return ""
        # 1. Extract inner text from [text]{u/upper/sub/...}
        t = re.sub(r'\[(.*?)\]\{(?:u|b|i|upper|sub|[a-zA-Z0-9#:,]+)\}', r'\1', text)
        # 2. Strip markdown bold/italic asterisks
        t = re.sub(r'\*\*(.*?)\*\*', r'\1', t)
        t = re.sub(r'\*(.*?)\*', r'\1', t)
        # 3. Replace <blank> / [BLANK] with underscores
        t = re.sub(r'<(?:blank|BLANK)>|\[(?:blank|BLANK)\]', '___________', t)
        return t.strip()

    def optimize_word_bank_layout(self, doc, words: List[str], printable_width_pt: float) -> Tuple[int, List[List[str]], float, List[float]]:
        """
        Optimally arranges Word Bank items across columns and rows to produce the most
        compact, balanced, and aesthetically pleasing box width and height.
        Evaluates column-major alphabetical, length-balanced bin-packing, and row-major permutations.
        """
        N = len(words)
        if N == 0:
            return 1, [[]], 50.0, [0.0]

        font_name = getattr(self, "font_name", "Times New Roman")
        font_size = getattr(self, "font_size", 12.0)

        clean_words = [self.strip_markup_for_measurement(w) for w in words]
        item_widths_pt = [self.measure_text_width_pt(doc, cw, font_name, font_size, is_bold=True) * 1.09 for cw in clean_words]
        max_item_w_pt = max(item_widths_pt) if item_widths_pt else 45.0

        pad_horiz_pt = cm_to_pt(0.20)      # 2.0 mm padding
        extra_buffer_pt = cm_to_pt(0.20)   # 2.0 mm corner buffer
        gap_pt = cm_to_pt(0.50)            # 5.0 mm inter-column gap

        avg_len = sum(len(cw) for cw in clean_words) / len(clean_words) if clean_words else 0

        # If items are long sentences / dialogue options (A. ... | B. ...), use 1 column with clean tight width
        is_sentence_options = any(re.match(r'^[A-Z]\.\s+', cw) for cw in clean_words) or (max_item_w_pt >= (printable_width_pt * 0.55)) or (avg_len >= 38)
        if is_sentence_options:
            needed_w = max_item_w_pt + (2 * pad_horiz_pt) + extra_buffer_pt
            box_w = min(printable_width_pt, needed_w)
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

            # 1. Alphabetical Column-Major
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

            # 3. Alphabetical Row-Major
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

        font_name = getattr(self, "font_name", "Times New Roman")
        font_size = getattr(self, "font_size", 12.0)

        is_word_bank = ('|' in raw_content) or (block.tag == "WORDBANK") or (block.tag.endswith(":bank"))

        is_adjacent_to_ins = False
        if blocks and 0 <= idx_block < len(blocks):
            if idx_block > 0 and blocks[idx_block - 1].tag == "INS":
                is_adjacent_to_ins = True
            elif idx_block + 1 < len(blocks) and blocks[idx_block + 1].tag == "INS":
                is_adjacent_to_ins = True
        elif getattr(self, "last_rendered_tag", None) == "INS":
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

            clean_lines = [self.strip_markup_for_measurement(l) for l in lines]
            line_widths_pt = [self.measure_text_width_pt(doc, cl, font_name, font_size, is_bold=True) * 1.11 for cl in clean_lines]
            max_line_w_pt = max(line_widths_pt) if line_widths_pt else 50.0

            pad_left_pt = cm_to_pt(0.20)     # Exactly 2.0 mm left margin
            pad_right_pt = cm_to_pt(0.20)    # Exactly 2.0 mm right margin
            extra_buffer_pt = cm_to_pt(0.20) # 2.0 mm corner clearance buffer
            total_pad_pt = pad_left_pt + pad_right_pt + extra_buffer_pt

            needed_w = max_line_w_pt + total_pad_pt
            box_width_pt = min(printable_width_pt, needed_w)

            num_lines = len(lines)
            box_height_pt = (num_lines * (font_size * 1.25)) + (font_size * 0.40) + 2.0

            try:
                shape = doc.Shapes.AddShape(5, 0, 0, box_width_pt, box_height_pt)
                tf = shape.TextFrame
                tf.MarginTop = 0
                tf.MarginBottom = 0
                tf.MarginLeft = pad_left_pt
                tf.MarginRight = pad_right_pt
                try:
                    tf.VerticalAnchor = 3  # msoAnchorMiddle (Vertical Center)
                except Exception:
                    pass
                try:
                    tf.WordWrap = -1
                except Exception:
                    pass
                try:
                    tf.AutoSize = False
                except Exception:
                    pass

                shape.Fill.Visible = False
                shape.Line.Weight = 1.0
                shape.Line.ForeColor.RGB = 0

                tr = tf.TextRange
                tr.Font.Name = font_name
                tr.Font.Size = font_size
                tr.Font.Bold = 1
                tr.Font.Color = 0
                tr.Text = "\n".join(self.clean_box_item(l) for l in lines)

                is_single_line = (num_lines == 1)
                for p_idx in range(1, tr.Paragraphs.Count + 1):
                    pf = tr.Paragraphs(p_idx).Range.ParagraphFormat
                    pf.SpaceBefore = 0
                    pf.SpaceAfter = 0
                    pf.LineSpacingRule = 0
                    pf.Alignment = 1 if is_single_line else 0

                try:
                    actual_lines = tr.ComputeStatistics(1)
                except Exception:
                    actual_lines = num_lines

                actual_box_height_pt = (actual_lines * (font_size * 1.25)) + (font_size * 0.40) + 2.0
                shape.Height = actual_box_height_pt

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
            pad_horiz_pt = cm_to_pt(0.20)
            extra_buffer_pt = cm_to_pt(0.25)
            left_offset_pt = max(0.0, (printable_width_pt - box_width_pt) / 2.0)

            num_rows = len(lines_bank)
            box_height_pt = (num_rows * (font_size * 1.25)) + (font_size * 0.40) + 2.0

            try:
                shape = doc.Shapes.AddShape(5, 0, 0, box_width_pt, box_height_pt)
                tf = shape.TextFrame
                tf.MarginTop = 0
                tf.MarginBottom = 0
                tf.MarginLeft = pad_horiz_pt
                tf.MarginRight = pad_horiz_pt
                try:
                    tf.VerticalAnchor = 3  # msoAnchorMiddle (Vertical Center)
                except Exception:
                    pass
                try:
                    tf.WordWrap = -1 if cols == 1 else 0
                except Exception:
                    pass
                try:
                    tf.AutoSize = False
                except Exception:
                    pass

                shape.Fill.Visible = False
                shape.Line.Weight = 1.0
                shape.Line.ForeColor.RGB = 0

                lines_text = ["\t".join(self.clean_box_item(it) for it in chunk) for chunk in lines_bank]
                tr = tf.TextRange
                tr.Font.Name = font_name
                tr.Font.Size = font_size
                tr.Font.Bold = 1
                tr.Font.Color = 0
                tr.Text = "\n".join(lines_text)

                is_single_item = (cols == 1 and num_rows == 1)
                for p_idx in range(1, tr.Paragraphs.Count + 1):
                    pf = tr.Paragraphs(p_idx).Range.ParagraphFormat
                    pf.SpaceBefore = 0
                    pf.SpaceAfter = 0
                    pf.LineSpacingRule = 0
                    pf.Alignment = 1 if is_single_item else 0
                    pf.TabStops.ClearAll()
                    for t_pt in tab_stops_pt[1:]:
                        pf.TabStops.Add(Position=t_pt, Alignment=0)

                try:
                    actual_lines = tr.ComputeStatistics(1)
                except Exception:
                    actual_lines = num_rows

                actual_box_height_pt = (actual_lines * (font_size * 1.25)) + (font_size * 0.40) + 2.0
                shape.Height = actual_box_height_pt

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
