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
    Parses color names ('red', 'blue', 'purple'), combo dropdown strings ('Blue #2563eb', 'Red #dc2626'),
    or hex strings ('#FF0000', '#003399', '#2563eb') into Windows COM BGR/RGB integer format: B * 65536 + G * 256 + R.
    """
    if not color_str:
        return None
    raw = str(color_str).strip()

    # 1. Search for 6-digit hex code anywhere in string (e.g. "Blue #2563eb", "Default (Black) #000000", "#2563eb")
    hex_match = re.search(r'#([0-9a-fA-F]{6})', raw)
    if hex_match:
        hex_val = hex_match.group(1).lower()
        r = int(hex_val[0:2], 16)
        g = int(hex_val[2:4], 16)
        b = int(hex_val[4:6], 16)
        return r + (g << 8) + (b << 16)

    # 2. Search for 3-digit hex code (#F00)
    hex3_match = re.search(r'#([0-9a-fA-F]{3})\b', raw)
    if hex3_match:
        h3 = hex3_match.group(1).lower()
        r = int(h3[0] * 2, 16)
        g = int(h3[1] * 2, 16)
        b = int(h3[2] * 2, 16)
        return r + (g << 8) + (b << 16)

    # 3. Direct color name lookup
    c = raw.lower()
    for name, val in COLOR_NAME_TO_RGB.items():
        if name in c:
            return val

    return None

def parse_highlight_to_index(hl_str: Optional[str]) -> Optional[int]:
    """Parses highlight name to MS Word WdColorIndex integer."""
    if not hl_str:
        return None
    h = hl_str.strip().lower()
    return HIGHLIGHT_NAME_TO_INDEX.get(h, None)
