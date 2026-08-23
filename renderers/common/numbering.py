import re
from typing import Optional
from uln_parser import ULNBlock
from renderers.common.units_and_colors import parse_color_to_rgb_int


class NumberingMixin:
    """Handles native Word list numbering and question number formatting."""

    def get_effective_number_format(self, extracted_pref: Optional[str], extracted_delim: Optional[str]) -> str:
        """
        Determines the effective list NumberFormat string:
        - Ensures a proper single space between prefix word and %1 (e.g. 'Question ' -> 'Question %1.').
        - For parenthesis prefixes (e.g. '(' -> '(%1)'), preserves tight bracket fit without unwanted space.
        - If delimiter was selected in GUI settings (and is non-default, e.g. ':', ')', '-'), use GUI delimiter.
          Otherwise use extracted delimiter from text or global default.
        """
        q_pref = getattr(self, "question_prefix", "")
        q_delim = getattr(self, "question_delimiter", ".")

        raw_pref = extracted_pref if (extracted_pref and extracted_pref.strip()) else (q_pref or "")
        if raw_pref and raw_pref.strip():
            p_strip = raw_pref.strip()
            if p_strip.endswith("("):
                pref = p_strip
            else:
                pref = f"{p_strip} "
        else:
            pref = ""

        if q_delim and q_delim != "." and "(" not in pref:
            delim = q_delim
        else:
            delim = extracted_delim if (extracted_delim and extracted_delim.strip()) else (q_delim or ".")
        return f"{pref}%1{delim}"

    def apply_native_numbered_list(self, word, sel, q_num: Optional[str] = None, number_format: Optional[str] = None):
        """
        Applies native MS Word Numbered List with guaranteed document-scoped font color & styling:
        - List number is ALWAYS BOLD by default
        - Text and numbering are separated by a SINGLE SPACE (TrailingCharacter = 1), NOT a tab!
        - LeftIndent and FirstLineIndent are flush at 0.0 cm (left page border)
        - Uses document-scoped ListTemplate for 100% reliable font color application across Word versions.
        - Starts a new independent list instance (ContinuePreviousList=False) on first question (q_num == '1' or is_first_question_in_num_block),
          and continues sequential list incrementing (ContinuePreviousList=True) on subsequent questions.
        """
        restart = getattr(self, "is_first_question_in_num_block", False) or (q_num == "1")
        self.is_first_question_in_num_block = False
        target_fmt = number_format if number_format else "%1."
        q_color_int = parse_color_to_rgb_int(getattr(self, "question_color", None))
        active_fmt = getattr(self, "_active_doc_list_template_fmt", None)

        try:
            doc = sel.Document
            # If restarting, format changed, or no active doc template exists, create a new document-scoped ListTemplate
            if restart or (active_fmt != target_fmt) or not hasattr(self, "_active_doc_list_template") or self._active_doc_list_template is None:
                list_tpl = doc.ListTemplates.Add(OutlineNumbered=False)
                lvl = list_tpl.ListLevels(1)
                lvl.TrailingCharacter = 1  # wdTrailingSpace = 1
                lvl.Font.Bold = 1          # ALWAYS BOLD
                lvl.Font.Name = getattr(self, "font_name", "Times New Roman")
                lvl.Font.Size = getattr(self, "font_size", 12.0)
                if q_color_int is not None:
                    lvl.Font.Color = q_color_int
                else:
                    lvl.Font.Color = 0
                lvl.NumberPosition = 0
                lvl.TextPosition = 0
                lvl.NumberFormat = target_fmt
                self._active_doc_list_template = list_tpl
                self._active_doc_list_template_fmt = target_fmt
                sel.Range.ListFormat.ApplyListTemplate(list_tpl, ContinuePreviousList=False)
            else:
                sel.Range.ListFormat.ApplyListTemplate(self._active_doc_list_template, ContinuePreviousList=True)

            sel.ParagraphFormat.LeftIndent = 0
            sel.ParagraphFormat.FirstLineIndent = 0
            sel.Font.Bold = 0
            sel.Font.Color = 0  # Black body text

        except Exception:
            try:
                list_tpl = word.ListGalleries(2).ListTemplates(1)
                lvl = list_tpl.ListLevels(1)
                lvl.TrailingCharacter = 1
                lvl.Font.Bold = 1
                lvl.NumberFormat = target_fmt
                if q_color_int is not None:
                    lvl.Font.Color = q_color_int
                sel.Range.ListFormat.ApplyListTemplate(list_tpl, ContinuePreviousList=not restart)
                sel.ParagraphFormat.LeftIndent = 0
                sel.ParagraphFormat.FirstLineIndent = 0
                sel.Font.Bold = 0
                sel.Font.Color = 0
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
