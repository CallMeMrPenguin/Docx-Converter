import os
import re
import hashlib
import tempfile
from typing import List, Optional, Tuple

SUPPORTED_IMAGE_EXTENSIONS = (
    '.png', '.jpg', '.jpeg', '.avif', '.avifs', '.webp', '.heic', '.heif',
    '.jfif', '.bmp', '.dib', '.gif', '.tiff', '.tif', '.ico', '.svg', '.wmf', '.emf'
)

def natural_sort_key(s: str) -> list:
    """
    Returns alphanumeric sorting key for natural sorting (e.g. img1, img2, img10 instead of img1, img10, img2).
    """
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

_pillow_plugins_registered = False

def register_image_plugins():
    global _pillow_plugins_registered
    if _pillow_plugins_registered:
        return
    try:
        import pillow_avif  # noqa: F401
    except Exception:
        pass
    try:
        import pillow_heif
        pillow_heif.register_heif_opener()
    except Exception:
        pass
    _pillow_plugins_registered = True

def ensure_word_compatible_image(image_path: str) -> Optional[str]:
    """
    Ensures an image is in a format natively supported by MS Word COM (PNG/JPEG/BMP).
    If the image is AVIF, WebP, HEIC/HEIF, TIFF, ICO, or any non-standard format,
    it automatically converts it to a high-quality temporary PNG file and returns the converted PNG path.
    """
    if not image_path or not os.path.exists(image_path):
        return None

    ext = os.path.splitext(image_path)[1].lower()
    native_word_exts = {'.png', '.jpg', '.jpeg', '.bmp', '.gif'}

    if ext in native_word_exts:
        return os.path.abspath(image_path)

    register_image_plugins()
    try:
        from PIL import Image

        stat = os.stat(image_path)
        cache_key = hashlib.md5(f"{os.path.abspath(image_path)}_{stat.st_mtime}_{stat.st_size}".encode('utf-8')).hexdigest()
        temp_dir = os.path.join(tempfile.gettempdir(), "docx_converter_img_cache")
        os.makedirs(temp_dir, exist_ok=True)
        cached_png = os.path.join(temp_dir, f"{cache_key}.png")

        if os.path.exists(cached_png) and os.path.getsize(cached_png) > 0:
            return cached_png

        with Image.open(image_path) as img:
            if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                conv_img = img.convert('RGBA')
            else:
                conv_img = img.convert('RGB')
            conv_img.save(cached_png, format='PNG')

        return cached_png
    except Exception as e:
        print(f"[ULNRenderer] Warning converting image {image_path} to PNG: {e}")
        return os.path.abspath(image_path)

def cm_to_pt(cm: float) -> float:
    return float(cm) * 28.346456692913385

def pt_to_cm(pt: float) -> float:
    return float(pt) / 28.346456692913385

COLOR_NAME_TO_RGB = {
    "red": 255,                 # RGB(255, 0, 0)
    "blue": 16711680,           # RGB(0, 0, 255)
    "green": 32768,             # RGB(0, 128, 0)
    "yellow": 65535,            # RGB(255, 255, 0)
    "purple": 8388736,          # RGB(128, 0, 128)
    "orange": 42495,            # RGB(255, 165, 0)
    "black": 0,
    "white": 16777215,
    "darkblue": 9125196,        # RGB(12, 15, 139)
    "grey": 8421504,
    "gray": 8421504,
}

HIGHLIGHT_NAME_TO_INDEX = {
    "yellow": 7,    # wdYellow = 7
    "green": 4,     # wdGreen = 4
    "cyan": 3,      # wdTurquoise = 3
    "pink": 5,      # wdPink = 5
    "blue": 9,      # wdBlue = 9
    "red": 6,       # wdRed = 6
    "gray": 15,     # wdGray25 = 15
}

def parse_color_to_rgb_int(color_str: str) -> Optional[int]:
    if not color_str:
        return None
    clean = color_str.strip().lower()
    hex_match = re.search(r'#([0-9a-fA-F]{6})', color_str)
    if hex_match:
        hex_val = hex_match.group(1).lower()
        r = int(hex_val[0:2], 16)
        g = int(hex_val[2:4], 16)
        b = int(hex_val[4:6], 16)
        return r + (g * 256) + (b * 65536)
    if clean in COLOR_NAME_TO_RGB:
        return COLOR_NAME_TO_RGB[clean]
    return None

def extract_question_prefix_and_body(text: str) -> Tuple[Optional[str], Optional[str], Optional[str], str]:
    """
    Universally extracts (full_prefix, delim_char, q_num, body_text) from any question text.
    Handles:
    - Leading markdown/ULN/HTML tags: [ins], [b], [i], [u], [hl], [ans], [q], [opt], [box], <b>, <strong>, <em>, **, *, _, etc.
    - Explicit '#' with optional space: #1., # 1., #1:, #1), #1/, #1-, #1 [PIC], (#1), [#1], (#1.), (#1:)
    - Word prefixes: Question #1., Câu #1:, Task #1., Activity #1., Ex #1., Exercise #1., Bài #1:, etc.
    - Standard numbering: 1., 1), 1:, 1/, 1 -, (1), [1]
    - Clean body extraction with balanced formatting tags.
    Returns (full_prefix, delim_char, q_num, body_text) if matched, else (None, None, None, original_text).
    """
    if not text or not text.strip():
        return None, None, None, text

    s = text.strip()

    pattern = re.compile(
        r'^\s*'
        r'((?:\[(?:ins|b|i|u|hl|ans|q|opt|box|p0|p1|p2|gap)\]|<(?:b|i|u|strong|em|span)[^>]*>|\*\*|\*|_|\{u\})*)'
        r'((?:Question|Câu|Task|Activity|Ex|Exercise|Part|Section|Item|Sentence|Dialogue|Problem|Bài)\s+)?'
        r'(?:(\(|\[)\s*)?'
        r'#?\s*(\d+)'
        r'([\.\:\)\/\-]|(?:\.\)|\:\)|\.\/|\:\-))?'
        r'(?:\s*(\)|\]))?'
        r'([\.\:\)\/\-])?'
        r'((?:\[\/(?:ins|b|i|u|hl|ans|q|opt|box|p0|p1|p2|gap)\]|<\/(?:b|i|u|strong|em|span)>|\*\*|\*|_|\{\/u\})*)'
        r'[\s\:\.\-]*'
        r'(.*)$',
        re.IGNORECASE | re.DOTALL
    )

    m = pattern.match(s)
    if not m:
        return None, None, None, text

    lead_tags = m.group(1) or ''
    prefix_word = m.group(2).strip() if m.group(2) else ''
    open_paren = m.group(3) or ''
    q_num = m.group(4)
    inner_delim = m.group(5) or ''
    close_paren = m.group(6) or ''
    outer_delim = m.group(7) or ''
    trail_tags = m.group(8) or ''
    raw_body = m.group(9).strip()

    # Safety check: If no explicit '#' was in the original text prefix, ensure it has either
    # a prefix word (Question, Câu, etc.), or parentheses ((1)), or a standard delimiter (., :, ), /, -)
    has_hash = ('#' in text[:m.start(9) if raw_body else len(text)])
    has_prefix_word = bool(prefix_word)
    has_paren = bool(open_paren and close_paren)
    has_delim = bool(inner_delim or outer_delim)

    if not (has_hash or has_prefix_word or has_paren or has_delim):
        return None, None, None, text

    # Balance/clean body text with formatting tags
    body = raw_body
    if body:
        # Check if trailing markdown tags exist in body that match lead_tags
        if '**' in lead_tags and re.search(r'\*\*(?:\[\/(?:ins|b|i|u|hl)\])*$', body) and not re.match(r'^(?:\[(?:ins|b|i|u|hl)\])*\*\*', body):
            m_lead = re.match(r'^(?:\[(?:ins|b|i|u|hl)\])+', body, re.IGNORECASE)
            if m_lead:
                body = body[:m_lead.end()] + '**' + body[m_lead.end():]
            else:
                body = '**' + body

        if '*' in lead_tags and not ('**' in lead_tags) and re.search(r'\*(?:\[\/(?:ins|b|i|u|hl)\])*$', body) and not re.match(r'^(?:\[(?:ins|b|i|u|hl)\])*\*', body):
            m_lead = re.match(r'^(?:\[(?:ins|b|i|u|hl)\])+', body, re.IGNORECASE)
            if m_lead:
                body = body[:m_lead.end()] + '*' + body[m_lead.end():]
            else:
                body = '*' + body

        for tag in ['ins', 'b', 'i', 'u', 'hl']:
            open_t = f'[{tag}]'
            close_t = f'[/{tag}]'
            if open_t in lead_tags.lower() and close_t in body.lower() and not body.lower().startswith(open_t):
                body = open_t + body

        # Clean orphan leading punctuation
        body = re.sub(r'^(?:[:\.\-\)])\s*', '', body).strip()

    if open_paren == '(' and close_paren == ')':
        full_prefix = f'{prefix_word} (' if prefix_word else '('
        delim_char = f'{inner_delim}){outer_delim}' if outer_delim else (f'{inner_delim})' if inner_delim else ')')
        return full_prefix, delim_char, q_num, body
    elif open_paren == '[' and close_paren == ']':
        full_prefix = f'{prefix_word} [' if prefix_word else '['
        delim_char = f'{inner_delim}]{outer_delim}' if outer_delim else (f'{inner_delim}]' if inner_delim else ']')
        return full_prefix, delim_char, q_num, body
    else:
        raw_delim = outer_delim or inner_delim
        if ':' in raw_delim:
            delim_char = ':'
        elif '.' in raw_delim:
            delim_char = '.'
        elif ')' in raw_delim:
            delim_char = ')'
        elif '/' in raw_delim:
            delim_char = '/'
        elif '-' in raw_delim:
            delim_char = '-'
        else:
            delim_char = '.'

        full_prefix = f'{prefix_word} ' if prefix_word else ''
        return full_prefix, delim_char, q_num, body

def split_line_into_option_items(line_text: str) -> List[str]:
    """
    Splits line text into multi-column items automatically by detecting:
    1. Explicit \\t characters
    2. Multiple sequential option choices like A. text B. text C. text D. text or a. b. c. d. or 1. 2. 3. 4.
    """
    if not line_text:
        return []

    if '\t' in line_text:
        return [x.strip() for x in line_text.split('\t') if x.strip()]

    # Extract potential question number prefix (e.g. "#1. ", "Question #1: ")
    pref, delim, q_num, rest_text = extract_question_prefix_and_body(line_text)
    q_prefix = f"{pref}#{q_num}{delim} " if q_num else ""

    # Search for option choices like A. B. C. D. or a. b. c. d. or (A) (B) (C) or 1. 2. 3. 4.
    pattern = r'(?:^|\s+)(?:(?:\*\*|\*|\[|\(?)*([a-zA-Z0-9][\.\)])(?:\*\*|\*|\]|\}|\{u\}|\))*)(?=\s+|$)'
    matches = list(re.finditer(pattern, rest_text))
    if len(matches) >= 2:
        labels = [re.sub(r'[\*\_\`\[\]\(\)\.\s]', '', m.group(1)) for m in matches]
        
        is_seq = False
        if all(len(l) == 1 and l.isalpha() for l in labels):
            if labels[0] == 'A' and all(ord(l) == ord('A') + i for i, l in enumerate(labels)):
                is_seq = True
            elif labels[0] == 'a' and all(ord(l) == ord('a') + i for i, l in enumerate(labels)):
                is_seq = True
        elif all(l.isdigit() for l in labels):
            nums = [int(l) for l in labels]
            if nums[0] == 1 and all(n == 1 + i for i, n in enumerate(nums)):
                is_seq = True

        if is_seq:
            items = []
            for i in range(len(matches)):
                start = matches[i].start()
                end = matches[i + 1].start() if i + 1 < len(matches) else len(rest_text)
                item_txt = rest_text[start:end].strip()
                if i == 0 and q_prefix:
                    item_txt = f"{q_prefix.strip()} {item_txt}"
                items.append(item_txt)
            return items

    return [line_text.strip()]

def apply_title_case_to_text(text: str) -> str:
    """Capitalizes first letter of each main word, keeping minor words in lowercase and preserving Roman numerals."""
    if not text:
        return text
    minor_words = {'a', 'an', 'the', 'and', 'but', 'or', 'for', 'nor', 'on', 'at', 'to', 'from', 'by', 'with', 'in', 'of', 'off'}
    words = text.split(' ')
    cased_words = []
    for idx, w in enumerate(words):
        if not w:
            cased_words.append(w)
            continue
        m = re.match(r'^([^\w]*)([\w\'-]+)([^\w]*)$', w)
        if not m:
            cased_words.append(w)
            continue
        pre, core, post = m.group(1), m.group(2), m.group(3)
        if re.match(r'^(?:[IVXLCDM]+|[A-Z]|\d+)$', core, re.IGNORECASE) and len(core) <= 4 and core.upper() in ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII", "A", "B", "C", "D", "E", "F", "G"]:
            cased_words.append(f"{pre}{core.upper()}{post}")
        elif idx == 0 or idx == len(words) - 1 or core.lower() not in minor_words:
            cased_words.append(f"{pre}{core.capitalize()}{post}")
        else:
            cased_words.append(f"{pre}{core.lower()}{post}")
    return ' '.join(cased_words)

def apply_sentence_case_to_text(text: str) -> str:
    """Capitalizes ONLY the first letter of the sentence, preserving Roman numeral / section letter prefix."""
    if not text:
        return text
    m = re.match(r'^(\s*(?:[IVXLCDM]+\.?|[A-Z]\.|\#?\d+[\.\:]?)\s*)(.*)$', text)
    if m:
        pref = m.group(1)
        body = m.group(2).strip()
        if body:
            body_cased = body[0].upper() + body[1:].lower()
            return f"{pref}{body_cased}"
        return text
    clean = text.strip()
    if clean:
        return clean[0].upper() + clean[1:].lower()
    return text


_gdi_measurer_instance = None


class GdiTextMeasurer:
    """Accurately measures physical text width in points using Windows GDI typography engine with caching."""

    def __init__(self):
        self._user32 = None
        self._gdi32 = None
        self._hdc = None
        self._dpi_x = 96
        self._font_cache = {}
        self._text_cache = {}
        self._init_gdi()

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

