import re
from typing import Optional

def cm_to_pt(cm: float) -> float:
    """Converts centimeters to typographical points (1 cm = 28.3465 pt)."""
    return float(cm) * 28.346456692913385

def pt_to_cm(pt: float) -> float:
    """Converts typographical points to centimeters."""
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
    "darkblue": 9,
    "gray": 16,     # wdGray25 = 16
    "grey": 16,
}

def parse_color_to_rgb_int(color_str: Optional[str]) -> Optional[int]:
    """
    Parses color names ('red', 'blue', 'purple') or hex strings ('#FF0000', '#003399')
    into Windows COM BGR/RGB integer format: B * 65536 + G * 256 + R.
    """
    if not color_str:
        return None
    c = color_str.strip().lower()
    if c in COLOR_NAME_TO_RGB:
        return COLOR_NAME_TO_RGB[c]
    if c.startswith('#') and len(c) in (4, 7):
        try:
            if len(c) == 4:
                r = int(c[1] * 2, 16)
                g = int(c[2] * 2, 16)
                b = int(c[3] * 2, 16)
            else:
                r = int(c[1:3], 16)
                g = int(c[3:5], 16)
                b = int(c[5:7], 16)
            return r + (g << 8) + (b << 16)
        except Exception:
            return None
    return None

def parse_highlight_to_index(hl_str: Optional[str]) -> Optional[int]:
    """Parses highlight name to MS Word WdColorIndex integer."""
    if not hl_str:
        return None
    h = hl_str.strip().lower()
    return HIGHLIGHT_NAME_TO_INDEX.get(h, None)
