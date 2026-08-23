import os
import re
from typing import Optional, List
from uln_parser import ULNBlock, parse_inline_spans, PicInfo, parse_pic_tag
from renderers.common.units_and_colors import cm_to_pt


class TableRendererMixin:
    """Renders bordered data grid tables ([TABLE]) and floating side-diagram MCQs."""

    def render_table(self, sel, doc, block_or_tdata, printable_width_cm: float, idx_block: int = 0, blocks: Optional[List[ULNBlock]] = None, **kwargs):
        """
        Renders native Microsoft Word Table objects for structured grid data,
        or uses paragraph tab stops for borderless lists.
        """
        tdata = block_or_tdata.table_data if hasattr(block_or_tdata, "table_data") else block_or_tdata
        if not tdata or not tdata.rows:
            return

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

        font_name = getattr(self, "font_name", "Times New Roman")
        font_size = getattr(self, "font_size", 12.0)

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

                last_tag = getattr(self, "last_rendered_tag", None)
                sel.ParagraphFormat.SpaceBefore = 14 if (idx_r == 0 and last_tag == "BOX") else 3
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

            if has_pic and pic_info_found and hasattr(self, "get_next_image_path"):
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
        last_tag = getattr(self, "last_rendered_tag", None)
        try:
            p_table_anchor.ParagraphFormat.SpaceBefore = 14.0 if (last_tag == "BOX") else 8.0
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
            tbl.TopPadding = cm_to_pt(0.15)
            tbl.BottomPadding = cm_to_pt(0.15)
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
            try:
                tbl.Rows(r_idx).HeightRule = 1  # wdRowHeightAtLeast = 1
                tbl.Rows(r_idx).Height = cm_to_pt(0.70)
            except Exception:
                pass

            for c_idx, cell_obj in enumerate(row_obj.cells, 1):
                if c_idx > num_cols:
                    break
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
                cell_range.Font.Name = font_name
                cell_range.Font.Size = font_size
                cell_range.ParagraphFormat.SpaceBefore = 0
                cell_range.ParagraphFormat.SpaceAfter = 0
                cell_range.ParagraphFormat.LineSpacingRule = 0
                cell_range.ParagraphFormat.Alignment = 1 if cell_obj.is_header else 0

                cell_txt = cell_obj.content.strip()
                if re.match(r'^\s*(?:_{2,}|<blank>|\[BLANK\])\s*$', cell_txt, re.IGNORECASE):
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
            tbl_end = tbl.Range.End
            rng_after = doc.Range(tbl_end, tbl_end)
            rng_after.Select()
            sel = doc.Application.Selection
            sel.ParagraphFormat.LeftIndent = 0
            sel.ParagraphFormat.RightIndent = 0
            sel.ParagraphFormat.FirstLineIndent = 0
            sel.ParagraphFormat.SpaceBefore = 8
            sel.ParagraphFormat.SpaceAfter = 4
            sel.ParagraphFormat.Alignment = 0
            sel.TypeParagraph()
        except Exception:
            try:
                tbl.Select()
                sel = doc.Application.Selection
                sel.Collapse(0)  # wdCollapseEnd = 0
                sel.TypeParagraph()
            except Exception:
                pass

        self.last_rendered_tag = "TABLE"
