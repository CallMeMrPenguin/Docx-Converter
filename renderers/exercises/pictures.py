import os
import re
from typing import List
from uln_parser import ULNBlock, PicInfo, parse_pic_tag
from renderers.common.units_and_colors import cm_to_pt


class PicturesRendererMixin:
    """Renders pictures, diagram illustrations, and 4-column picture grids ([PIC_GRID])."""

    def render_pic(self, sel, doc, pic: PicInfo):
        """Renders an image file from user queue in order, or falls back to 'test pic/' folder."""
        get_next_path = getattr(self, "get_next_image_path", None)
        target_path = get_next_path(pic) if get_next_path else None

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
        font_name = getattr(self, "font_name", "Times New Roman")
        sel.Font.Name = font_name
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

    def render_pic_grid(self, sel, doc, children: List[ULNBlock], printable_width_cm: float):
        """
        Renders 4-column horizontal picture grid [PIC_GRID] using pure paragraph Tab Stops (no Table object).
        Each row has a picture line and an aligned caption line using identical Center Tab Stops.
        """
        if not children:
            return

        font_name = getattr(self, "font_name", "Times New Roman")
        font_size = getattr(self, "font_size", 12.0)

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
                get_next_path = getattr(self, "get_next_image_path", None)
                target_path = get_next_path(pic_info) if get_next_path else None

                if target_path and os.path.exists(target_path):
                    try:
                        col_w_pt = cm_to_pt(slot_w_cm)
                        shp = sel.InlineShapes.AddPicture(FileName=os.path.abspath(target_path))
                        shp.Width = min(col_w_pt - 10.0, cm_to_pt(3.6))
                        shp.Height = cm_to_pt(2.6)
                    except Exception as e:
                        print(f"[ULNRenderer] Warning in pic_grid picture: {e}")
                else:
                    sel.Font.Name = font_name
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
                sel.Font.Name = font_name
                sel.Font.Size = font_size
                sel.Font.Bold = 0
                sel.Font.Italic = 0
                sel.Font.Color = 0

                col_w_pt = cm_to_pt(slot_w_cm)
                img_w_pt = min(col_w_pt - 10.0, cm_to_pt(3.6))

                from renderer_utils import extract_question_prefix_and_body
                clean_content = re.sub(r'\[PIC(?::[^\]]*)?\]', '', child.content, flags=re.IGNORECASE).strip()
                _pref, _delim, q_num_ext, body_ext = extract_question_prefix_and_body(clean_content)
                num_part = q_num_ext if q_num_ext is not None else str(global_idx + 1)
                body_part = body_ext.strip() if body_ext else ""

                from renderers.common.units_and_colors import parse_color_to_rgb_int
                q_col_int = parse_color_to_rgb_int(getattr(self, "question_color", None))

                prefix_w_pt = len(f"{num_part}. ") * (font_size * 0.48)
                char_under_w_pt = max(4.0, font_size * 0.44)
                num_underscores = max(10, int((img_w_pt - prefix_w_pt) / char_under_w_pt))

                sel.Font.Name = font_name
                sel.Font.Size = font_size
                sel.Font.Bold = 1
                sel.Font.Color = q_col_int if q_col_int is not None else 0
                sel.TypeText(f"{num_part}. ")
                sel.Font.Bold = 0
                sel.Font.Color = 0

                if not body_part or "<blank>" in body_part.lower() or "[blank]" in body_part.lower() or "_" in body_part:
                    sel.TypeText("_" * num_underscores)
                else:
                    self.write_inline_spans(sel, parse_inline_spans(body_part))

            sel.TypeParagraph()

        sel.ParagraphFormat.TabStops.ClearAll()
        sel.ParagraphFormat.LeftIndent = 0
        sel.ParagraphFormat.RightIndent = 0
        sel.ParagraphFormat.FirstLineIndent = 0
        sel.ParagraphFormat.SpaceBefore = 6
        sel.ParagraphFormat.SpaceAfter = 4
        sel.ParagraphFormat.Alignment = 0

        self.last_rendered_tag = "PIC_GRID"
