"""
gui/theme.py - Authentic Flat Minimal Light UI Theme System.
Implements zero-border philosophy, using contrast surfaces, subtle slate dividers,
and clean typography for an elegant, distraction-free editing experience.
"""

import tkinter as tk
from tkinter import ttk

class Theme:
    # ── Canvas & Contrast Surfaces ──────────────────────────────
    BG_CANVAS = "#f8fafc"        # Slate 50: clean airy background
    BG_SURFACE = "#ffffff"       # Pure white: card/panel surfaces
    BG_SURFACE_ALT = "#f1f5f9"   # Slate 100: header, toolbar & soft inset
    BG_CARD = "#ffffff"          # Elevated container surface
    BG_HOVER = "#e2e8f0"         # Slate 200: subtle button & row hover
    BG_ACTIVE = "#cbd5e1"        # Slate 300: active press state

    # ── Text & Typography ───────────────────────────────────────
    TEXT_PRIMARY = "#0f172a"     # Slate 900: high-contrast text
    TEXT_SECONDARY = "#475569"   # Slate 600: labels & secondary info
    TEXT_MUTED = "#94a3b8"       # Slate 400: hints & placeholders
    TEXT_ON_PRIMARY = "#ffffff"  # White text on vibrant buttons

    # ── Brand Accents & State Colors ────────────────────────────
    PRIMARY = "#4f46e5"          # Indigo 600: primary brand accent
    PRIMARY_HOVER = "#4338ca"    # Indigo 700: primary hover
    PRIMARY_LIGHT = "#e0e7ff"    # Indigo 100: selected item highlight
    
    SUCCESS = "#059669"          # Emerald 600: compile / run action
    SUCCESS_HOVER = "#047857"    # Emerald 700: compile hover
    
    WARNING = "#d97706"          # Amber 600: replace action
    WARNING_HOVER = "#b45309"
    
    DANGER = "#e11d48"           # Rose 600: delete / clear action
    DANGER_HOVER = "#be123c"
    
    INFO = "#0284c7"             # Sky 600: import / docx action
    INFO_HOVER = "#0369a1"
    
    PURPLE = "#7c3aed"           # Violet 600: image cross-check
    PURPLE_HOVER = "#6d28d9"

    # ── Subtle Dividers (Zero Harsh Borders) ────────────────────
    DIVIDER = "#e2e8f0"          # Slate 200: 1px divider line

    # ── Font Definitions ────────────────────────────────────────
    FONT_TITLE = ("Segoe UI", 13, "bold")
    FONT_SUBTITLE = ("Segoe UI", 9)
    FONT_SECTION = ("Segoe UI", 10, "bold")
    FONT_BODY = ("Segoe UI", 9)
    FONT_BODY_BOLD = ("Segoe UI", 9, "bold")
    FONT_SMALL = ("Segoe UI", 8)
    FONT_SMALL_BOLD = ("Segoe UI", 8, "bold")
    FONT_CODE = ("Consolas", 10)
    FONT_CODE_EDITOR = ("Consolas", 11)


def apply_light_theme(root: tk.Tk) -> ttk.Style:
    """Configures modern Flat Minimal Light UI styling across all Tk and TTK widgets."""
    style = ttk.Style()
    style.theme_use("clam")

    # Global TTK defaults: zero border, clean contrast
    style.configure(".", background=Theme.BG_CANVAS, foreground=Theme.TEXT_PRIMARY, font=Theme.FONT_BODY)
    style.configure("TLabel", background=Theme.BG_CANVAS, foreground=Theme.TEXT_PRIMARY)
    style.configure("TFrame", background=Theme.BG_CANVAS)
    
    # Surface Cards (replacing bordered LabelFrames with clean white surfaces)
    style.configure("TLabelframe", background=Theme.BG_SURFACE, borderwidth=0, relief="flat")
    style.configure("TLabelframe.Label", background=Theme.BG_SURFACE, foreground=Theme.TEXT_PRIMARY, font=Theme.FONT_SECTION)

    # Flat TTK Buttons
    style.configure(
        "TButton",
        font=Theme.FONT_BODY_BOLD,
        background=Theme.BG_SURFACE_ALT,
        foreground=Theme.TEXT_PRIMARY,
        borderwidth=0,
        relief="flat",
        padding=(10, 5)
    )
    style.map(
        "TButton",
        background=[("active", Theme.BG_HOVER), ("pressed", Theme.BG_ACTIVE)],
        foreground=[("active", Theme.TEXT_PRIMARY)]
    )

    # Primary Action Button Style
    style.configure(
        "Primary.TButton",
        font=Theme.FONT_BODY_BOLD,
        background=Theme.PRIMARY,
        foreground=Theme.TEXT_ON_PRIMARY,
        borderwidth=0,
        relief="flat",
        padding=(12, 6)
    )
    style.map("Primary.TButton", background=[("active", Theme.PRIMARY_HOVER)])

    # Success Action Button Style (Compile)
    style.configure(
        "Success.TButton",
        font=Theme.FONT_BODY_BOLD,
        background=Theme.SUCCESS,
        foreground=Theme.TEXT_ON_PRIMARY,
        borderwidth=0,
        relief="flat",
        padding=(14, 6)
    )
    style.map("Success.TButton", background=[("active", Theme.SUCCESS_HOVER)])

    # Combobox: clean flat surface, zero border, indigo focus
    style.configure(
        "TCombobox",
        fieldbackground=Theme.BG_SURFACE,
        background=Theme.BG_SURFACE_ALT,
        foreground=Theme.TEXT_PRIMARY,
        selectbackground=Theme.PRIMARY_LIGHT,
        selectforeground=Theme.PRIMARY,
        arrowcolor=Theme.TEXT_SECONDARY,
        borderwidth=0,
        relief="flat",
        padding=4
    )
    style.map(
        "TCombobox",
        fieldbackground=[("readonly", Theme.BG_SURFACE), ("focus", Theme.BG_SURFACE)],
        foreground=[("readonly", Theme.TEXT_PRIMARY)]
    )

    # Spinbox: clean flat input
    style.configure(
        "TSpinbox",
        fieldbackground=Theme.BG_SURFACE,
        background=Theme.BG_SURFACE_ALT,
        foreground=Theme.TEXT_PRIMARY,
        selectbackground=Theme.PRIMARY_LIGHT,
        selectforeground=Theme.PRIMARY,
        arrowcolor=Theme.TEXT_SECONDARY,
        borderwidth=0,
        relief="flat",
        padding=4
    )

    # Combobox popup listbox styling
    root.option_add("*TCombobox*Listbox.background", Theme.BG_SURFACE)
    root.option_add("*TCombobox*Listbox.foreground", Theme.TEXT_PRIMARY)
    root.option_add("*TCombobox*Listbox.selectBackground", Theme.PRIMARY_LIGHT)
    root.option_add("*TCombobox*Listbox.selectForeground", Theme.PRIMARY)
    root.option_add("*TCombobox*Listbox.font", Theme.FONT_BODY)
    root.option_add("*TCombobox*Listbox.borderwidth", "0")
    root.option_add("*TCombobox*Listbox.relief", "flat")

    # Treeview: clean alternating rows, flat header, zero border
    style.configure(
        "Treeview",
        background=Theme.BG_SURFACE,
        foreground=Theme.TEXT_PRIMARY,
        fieldbackground=Theme.BG_SURFACE,
        rowheight=26,
        font=Theme.FONT_BODY,
        borderwidth=0,
        relief="flat"
    )
    style.map(
        "Treeview",
        background=[("selected", Theme.PRIMARY_LIGHT)],
        foreground=[("selected", Theme.PRIMARY)]
    )
    style.configure(
        "Treeview.Heading",
        background=Theme.BG_SURFACE_ALT,
        foreground=Theme.TEXT_SECONDARY,
        font=Theme.FONT_SMALL_BOLD,
        borderwidth=0,
        relief="flat"
    )
    style.map("Treeview.Heading", background=[("active", Theme.BG_HOVER)])

    # Scrollbars: subtle flat slate
    style.configure(
        "TScrollbar",
        background=Theme.BG_SURFACE_ALT,
        troughcolor=Theme.BG_CANVAS,
        borderwidth=0,
        arrowsize=12,
        relief="flat"
    )
    style.map("TScrollbar", background=[("active", Theme.BG_HOVER)])

    # Progress bar: Indigo flat bar
    style.configure(
        "Horizontal.TProgressbar",
        background=Theme.PRIMARY,
        troughcolor=Theme.BG_SURFACE_ALT,
        borderwidth=0,
        thickness=4
    )

    return style
