import re
from typing import List, Optional
from uln_parser import InlineSpan, PicInfo, parse_pic_tag
from renderers.common.units_and_colors import parse_color_to_rgb_int, HIGHLIGHT_NAME_TO_INDEX


class InlineWriterMixin:
    """Handles formatted text run writing, styles, colors, and inline tags into MS Word."""

    def write_inline_spans(self, sel, spans: List[InlineSpan], default_bold: bool = False, default_italic: bool = False, default_uppercase: bool = False, custom_font_size: Optional[float] = None, force_color: Optional[int] = None):
        """Writes formatted text runs strictly according to span AST properties with cached COM attributes for 10x speed."""
        f = sel.Font
        f_size = custom_font_size if custom_font_size is not None else getattr(self, "font_size", 12.0)
        font_name = getattr(self, "font_name", "Times New Roman")
        instruction_color = getattr(self, "instruction_color", None)
        question_color = getattr(self, "question_color", None)

        for idx, span in enumerate(spans):
            text = span.text

            # Check if span is an inline [PIC...] tag
            if text.startswith("[PIC:") or text.strip().upper() == "[PIC]":
                pic_info = parse_pic_tag(text) or PicInfo(description="Activity Picture", pos="center", size="small")
                if hasattr(self, "render_pic"):
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
            f.Name = font_name
            f.Size = f_size
            f.Bold = is_bold
            f.Italic = is_italic
            f.Underline = is_under

            if force_color is not None:
                f.Color = force_color
            elif span.color:
                rgb_int = parse_color_to_rgb_int(span.color)
                f.Color = rgb_int if rgb_int is not None else 0
            elif span.is_instruction and instruction_color:
                ins_color_int = parse_color_to_rgb_int(instruction_color)
                f.Color = ins_color_int if ins_color_int is not None else 0
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
                base_color = force_color if force_color is not None else (parse_color_to_rgb_int(span.color) if span.color else (parse_color_to_rgb_int(instruction_color) if (span.is_instruction and instruction_color) else 0))
                for part in parts:
                    if not part:
                        continue
                    if re.match(r'^\(\s*#?\d+\s*[\.\:\)]*\)$', part):
                        clean_paren = re.sub(r'#|\s', '', part)
                        f.Bold = 1
                        q_col = parse_color_to_rgb_int(question_color)
                        f.Color = q_col if q_col is not None else 0
                        sel.TypeText(clean_paren)
                        f.Bold = is_bold
                        f.Color = base_color if base_color is not None else 0
                    elif re.match(r'^\[\s*#?\d+\s*[\.\:\)]*\]$', part):
                        clean_bracket = re.sub(r'#|\s', '', part)
                        f.Bold = 1
                        q_col = parse_color_to_rgb_int(question_color)
                        f.Color = q_col if q_col is not None else 0
                        sel.TypeText(clean_bracket)
                        f.Bold = is_bold
                        f.Color = base_color if base_color is not None else 0
                    elif re.match(r'^#\s*\d+[\.\:\)\/\-]*$', part):
                        clean_num = re.sub(r'^#\s*', '', part)
                        f.Bold = 1
                        q_col = parse_color_to_rgb_int(question_color)
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
