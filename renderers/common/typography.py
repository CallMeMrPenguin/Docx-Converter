import re
from typing import List, Optional
from renderers.common.units_and_colors import cm_to_pt, pt_to_cm

_gdi_measurer_instance = None


def apply_font_family(font_obj, font_name: str, font_size: Optional[float] = None):
    """Sets font family uniformly across Ascii, Other, and Complex/Bidi to prevent font inconsistencies in Word."""
    if not font_obj or not font_name:
        return
    try:
        font_obj.Name = font_name
    except Exception:
        pass
    try:
        font_obj.NameAscii = font_name
    except Exception:
        pass
    try:
        font_obj.NameOther = font_name
    except Exception:
        pass
    try:
        font_obj.NameBi = font_name
    except Exception:
        pass
    if font_size is not None:
        try:
            font_obj.Size = font_size
        except Exception:
            pass


class GdiTextMeasurer:
    """Accurately measures physical text width in points using Windows GDI typography engine with caching."""

    def __init__(self):
        self._user32 = None
        self._gdi32 = None
        self._hdc = None
        self._dpi_x = 96
        self._font_cache = {}
        self._text_cache = {}
        self._max_cache_size = 10000
        self._init_gdi()

    def __del__(self):
        try:
            if self._gdi32 and self._font_cache:
                for hfont in self._font_cache.values():
                    self._gdi32.DeleteObject(hfont)
                self._font_cache.clear()
            if self._user32 and self._hdc:
                self._user32.ReleaseDC(0, self._hdc)
                self._hdc = None
        except Exception:
            pass

    def _init_gdi(self):
        try:
            import ctypes
            self._user32 = ctypes.windll.user32
            self._gdi32 = ctypes.windll.gdi32
            self._hdc = self._user32.GetDC(0)
            if self._hdc:
                self._dpi_x = self._gdi32.GetDeviceCaps(self._hdc, 88) or 96  # LOGPIXELSX = 88
        except Exception:
            self._user32 = None
            self._gdi32 = None
            self._hdc = None

    def measure_text_pt(self, text: str, font_name: str = "Times New Roman", font_size_pt: float = 12.0, is_bold: bool = False, is_italic: bool = False) -> float:
        if not text:
            return 0.0
        cache_key = (text, font_name, font_size_pt, is_bold, is_italic)
        if cache_key in self._text_cache:
            return self._text_cache[cache_key]

        if len(self._text_cache) >= self._max_cache_size:
            # Drop the oldest 2000 entries to prevent memory leak
            keys = list(self._text_cache.keys())[:2000]
            for k in keys:
                self._text_cache.pop(k, None)

        if self._gdi32 and self._hdc:
            try:
                import ctypes
                from ctypes import wintypes

                font_key = (font_name, font_size_pt, is_bold, is_italic)
                if font_key not in self._font_cache:
                    height_px = -int(round(font_size_pt * self._dpi_x / 72.0))
                    weight = 700 if is_bold else 400
                    italic = 1 if is_italic else 0
                    hfont = self._gdi32.CreateFontW(
                        height_px, 0, 0, 0,
                        weight, italic, 0, 0,
                        1, 0, 0, 0, 0,
                        font_name
                    )
                    self._font_cache[font_key] = hfont

                hfont = self._font_cache[font_key]
                old_font = self._gdi32.SelectObject(self._hdc, hfont)

                class SIZE(ctypes.Structure):
                    _fields_ = [("cx", wintypes.LONG), ("cy", wintypes.LONG)]

                size = SIZE()
                self._gdi32.GetTextExtentPoint32W(self._hdc, text, len(text), ctypes.byref(size))
                self._gdi32.SelectObject(self._hdc, old_font)

                width_pt = float(size.cx) * 72.0 / float(self._dpi_x)
                if width_pt > 0:
                    self._text_cache[cache_key] = width_pt
                    return width_pt
            except Exception:
                pass

        # Fallback estimation
        char_w = font_size_pt * (0.50 if is_bold else 0.45)
        calc_w = len(text) * char_w
        self._text_cache[cache_key] = calc_w
        return calc_w


def get_gdi_text_measurer() -> GdiTextMeasurer:
    global _gdi_measurer_instance
    if _gdi_measurer_instance is None:
        _gdi_measurer_instance = GdiTextMeasurer()
    return _gdi_measurer_instance


class TypographyMixin:
    """Provides shared text measurement, markup stripping, and visual line calculations."""

    def apply_font_family(self, font_obj, font_name: Optional[str] = None, font_size: Optional[float] = None):
        """Sets font family uniformly across Ascii, Other, and Complex/Bidi on any Word Font or Range object."""
        target_font = font_name or getattr(self, "font_name", "Times New Roman")
        apply_font_family(font_obj, target_font, font_size)

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
        t = re.sub(r'\[(.*?)\]\{(?:u|b|i|[a-zA-Z0-9#:,]+)\}', r'\1', t)
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

        font_name = getattr(self, "font_name", "Times New Roman")
        font_size = getattr(self, "font_size", 12.0)

        space_w = self.measure_text_width_pt(doc, " ", font_name, font_size, is_bold=is_bold)
        curr_line_w = 0.0
        lines_count = 1

        for w in words:
            w_pt = self.measure_text_width_pt(doc, w, font_name, font_size, is_bold=is_bold) * 1.02
            if curr_line_w == 0.0:
                curr_line_w = w_pt
            elif curr_line_w + space_w + w_pt <= avail_width_pt:
                curr_line_w += space_w + w_pt
            else:
                lines_count += 1
                curr_line_w = w_pt

        return max(1, lines_count)

    def wrap_text_into_lines(self, doc, text: str, max_w_pt: float, is_bold: bool = False) -> List[str]:
        """
        Splits a text string into wrapped lines that fit within max_w_pt using GDI text measurement.
        """
        if not text:
            return [""]
        words = text.strip().split()
        if not words:
            return [""]

        font_name = getattr(self, "font_name", "Times New Roman")
        font_size = getattr(self, "font_size", 12.0)

        space_w = self.measure_text_width_pt(doc, " ", font_name, font_size, is_bold=is_bold)
        lines = []
        curr_words = []
        curr_line_w = 0.0

        for w in words:
            w_clean = self.strip_markup_for_measurement(w)
            w_pt = self.measure_text_width_pt(doc, w_clean, font_name, font_size, is_bold=is_bold) * 1.04

            if not curr_words:
                curr_words.append(w)
                curr_line_w = w_pt
            elif curr_line_w + space_w + w_pt <= max_w_pt:
                curr_words.append(w)
                curr_line_w += space_w + w_pt
            else:
                lines.append(" ".join(curr_words))
                curr_words = [w]
                curr_line_w = w_pt

        if curr_words:
            lines.append(" ".join(curr_words))

        return lines if lines else [""]

