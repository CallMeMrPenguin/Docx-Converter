"""
gui/components/status_bar.py - Flat Minimal Status Bar Component.
Displays document statistics (lines, words, characters), compilation progress,
and a non-blocking cancellation button with zero borders.
"""

import tkinter as tk
from tkinter import ttk
from typing import Optional, Callable
from gui.theme import Theme


class StatusBar(tk.Frame):
    """Clean flat status bar for tracking document metrics and background task state."""

    def __init__(self, parent, on_cancel: Optional[Callable[[], None]] = None, **kwargs):
        super().__init__(parent, bg=Theme.BG_SURFACE_ALT, height=28, padx=12, pady=4, **kwargs)
        self.on_cancel_callback = on_cancel

        # ── Left: Status & Message ───────────────────────────────
        self.status_lbl = tk.Label(
            self,
            text="Sẵn sàng",
            font=Theme.FONT_SMALL,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.TEXT_SECONDARY
        )
        self.status_lbl.pack(side="left")

        # ── Progress Bar & Cancel Button (Hidden by default) ────
        self.progress_frame = tk.Frame(self, bg=Theme.BG_SURFACE_ALT)
        
        self.progress_bar = ttk.Progressbar(
            self.progress_frame,
            orient="horizontal",
            length=140,
            mode="determinate",
            style="Horizontal.TProgressbar"
        )
        self.progress_bar.pack(side="left", padx=(8, 6))

        self.btn_cancel = tk.Button(
            self.progress_frame,
            text="Hủy",
            command=self._trigger_cancel,
            bg=Theme.DANGER,
            fg=Theme.TEXT_ON_PRIMARY,
            activebackground=Theme.DANGER_HOVER,
            font=Theme.FONT_SMALL_BOLD,
            relief="flat",
            bd=0,
            padx=6,
            pady=1,
            cursor="hand2"
        )
        self.btn_cancel.pack(side="left")

        # ── Right: Stats Badges (Separated by Pipes) ─────────────
        stats_frame = tk.Frame(self, bg=Theme.BG_SURFACE_ALT)
        stats_frame.pack(side="right")

        self.lines_lbl = tk.Label(stats_frame, text="0 dòng", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE_ALT, fg=Theme.TEXT_MUTED)
        self.lines_lbl.pack(side="left", padx=4)

        tk.Label(stats_frame, text="|", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE_ALT, fg=Theme.DIVIDER).pack(side="left")

        self.words_lbl = tk.Label(stats_frame, text="0 từ", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE_ALT, fg=Theme.TEXT_MUTED)
        self.words_lbl.pack(side="left", padx=4)

        tk.Label(stats_frame, text="|", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE_ALT, fg=Theme.DIVIDER).pack(side="left")

        self.chars_lbl = tk.Label(stats_frame, text="0 ký tự", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE_ALT, fg=Theme.TEXT_MUTED)
        self.chars_lbl.pack(side="left", padx=4)

    def set_message(self, message: str, is_error: bool = False):
        color = Theme.DANGER if is_error else Theme.TEXT_SECONDARY
        self.status_lbl.config(text=message, fg=color)

    def update_stats(self, text: str):
        lines = text.count("\n") + (1 if text else 0)
        words = len(text.split())
        chars = len(text)
        self.lines_lbl.config(text=f"{lines} dòng")
        self.words_lbl.config(text=f"{words} từ")
        self.chars_lbl.config(text=f"{chars} ký tự")

    def show_progress(self, percent: float = 0.0, message: Optional[str] = None):
        self.progress_frame.pack(side="left", padx=(10, 0))
        self.progress_bar["value"] = percent * 100
        if message:
            self.set_message(message)

    def hide_progress(self):
        self.progress_frame.pack_forget()

    def _trigger_cancel(self):
        if self.on_cancel_callback:
            self.on_cancel_callback()
