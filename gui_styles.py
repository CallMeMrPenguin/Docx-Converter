"""
gui_styles.py - Backward-compatible facade for theme styling.
Points to gui.theme.apply_light_theme.
"""

from gui.theme import Theme, apply_light_theme

def setup_dark_theme(root):
    """Backward-compatible function name applying the modern Flat Minimal Light UI theme."""
    return apply_light_theme(root)

__all__ = ["setup_dark_theme", "apply_light_theme", "Theme"]
