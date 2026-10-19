"""
gui/components/settings_panel.py - Flat Minimal Settings Sidebar Component.
Organizes typography, margin, numbering, and color options on a clean contrast surface.
"""

import tkinter as tk
from tkinter import ttk
from typing import Dict, Any
from gui.theme import Theme


class SettingsPanel(tk.Frame):
    """Clean sidebar panel containing all document and typography styling controls."""

    COLOR_CHOICES = [
        "Default (Black) #000000",
        "Blue #2563eb",
        "Navy #1e40af",
        "Red #dc2626",
        "Emerald #059669",
        "Purple #7c3aed",
        "Amber #d97706",
        "Dark Grey #475569"
    ]

    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=Theme.BG_SURFACE, padx=14, pady=14, **kwargs)

        # ── Title Header ─────────────────────────────────────────
        lbl_head = tk.Label(
            self,
            text="Tùy chỉnh định dạng",
            font=Theme.FONT_SECTION,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_PRIMARY
        )
        lbl_head.pack(anchor="w", pady=(0, 10))

        # ── Typography Section ───────────────────────────────────
        tk.Label(self, text="Phông chữ:", font=Theme.FONT_SMALL_BOLD, bg=Theme.BG_SURFACE, fg=Theme.TEXT_SECONDARY).pack(anchor="w", pady=(2, 2))
        self.font_var = tk.StringVar(value="Times New Roman")
        font_cb = ttk.Combobox(
            self,
            textvariable=self.font_var,
            values=["Times New Roman", "Arial", "Calibri", "Cambria", "Georgia", "Garamond"],
            width=22
        )
        font_cb.pack(fill="x", pady=(0, 6))

        tk.Label(self, text="Cỡ chữ (pt):", font=Theme.FONT_SMALL_BOLD, bg=Theme.BG_SURFACE, fg=Theme.TEXT_SECONDARY).pack(anchor="w", pady=(2, 2))
        self.size_var = tk.DoubleVar(value=12.0)
        size_sp = ttk.Spinbox(self, from_=8.0, to=24.0, increment=0.5, textvariable=self.size_var, width=22)
        size_sp.pack(fill="x", pady=(0, 10))

        # ── Margins Section ──────────────────────────────────────
        tk.Label(self, text="Lề trang (cm):", font=Theme.FONT_SMALL_BOLD, bg=Theme.BG_SURFACE, fg=Theme.TEXT_SECONDARY).pack(anchor="w", pady=(4, 4))
        m_frame = tk.Frame(self, bg=Theme.BG_SURFACE)
        m_frame.pack(fill="x", pady=(0, 10))

        # Row 0: Top / Bottom
        tk.Label(m_frame, text="Trên:", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE, fg=Theme.TEXT_SECONDARY).grid(row=0, column=0, sticky="w", pady=2)
        self.m_top_var = tk.DoubleVar(value=2.0)
        tk.Spinbox(m_frame, from_=0.5, to=5.0, increment=0.5, textvariable=self.m_top_var, width=4, bg=Theme.BG_SURFACE_ALT, fg=Theme.TEXT_PRIMARY, relief="flat", highlightthickness=0).grid(row=0, column=1, padx=3, pady=2)

        tk.Label(m_frame, text="Dưới:", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE, fg=Theme.TEXT_SECONDARY).grid(row=0, column=2, sticky="w", pady=2, padx=(6, 0))
        self.m_bottom_var = tk.DoubleVar(value=2.0)
        tk.Spinbox(m_frame, from_=0.5, to=5.0, increment=0.5, textvariable=self.m_bottom_var, width=4, bg=Theme.BG_SURFACE_ALT, fg=Theme.TEXT_PRIMARY, relief="flat", highlightthickness=0).grid(row=0, column=3, padx=3, pady=2)

        # Row 1: Left / Right
        tk.Label(m_frame, text="Trái:", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE, fg=Theme.TEXT_SECONDARY).grid(row=1, column=0, sticky="w", pady=2)
        self.m_left_var = tk.DoubleVar(value=3.0)
        tk.Spinbox(m_frame, from_=0.5, to=5.0, increment=0.5, textvariable=self.m_left_var, width=4, bg=Theme.BG_SURFACE_ALT, fg=Theme.TEXT_PRIMARY, relief="flat", highlightthickness=0).grid(row=1, column=1, padx=3, pady=2)

        tk.Label(m_frame, text="Phải:", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE, fg=Theme.TEXT_SECONDARY).grid(row=1, column=2, sticky="w", pady=2, padx=(6, 0))
        self.m_right_var = tk.DoubleVar(value=1.5)
        tk.Spinbox(m_frame, from_=0.5, to=5.0, increment=0.5, textvariable=self.m_right_var, width=4, bg=Theme.BG_SURFACE_ALT, fg=Theme.TEXT_PRIMARY, relief="flat", highlightthickness=0).grid(row=1, column=3, padx=3, pady=2)

        # ── Question & Option Styling Section ───────────────────
        tk.Label(self, text="Kiểu câu hỏi & đáp án:", font=Theme.FONT_SMALL_BOLD, bg=Theme.BG_SURFACE, fg=Theme.TEXT_SECONDARY).pack(anchor="w", pady=(4, 4))
        q_frame = tk.Frame(self, bg=Theme.BG_SURFACE_ALT, padx=8, pady=6)
        q_frame.pack(fill="x", pady=(0, 10))

        tk.Label(q_frame, text="Tiền tố:", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE_ALT, fg=Theme.TEXT_SECONDARY).grid(row=0, column=0, sticky="w", pady=2)
        self.q_prefix_var = tk.StringVar(value="")
        ttk.Combobox(q_frame, textvariable=self.q_prefix_var, values=["", "Question ", "Câu ", "Task ", "Exercise ", "Ex "], width=14).grid(row=0, column=1, padx=4, pady=2, sticky="ew")

        tk.Label(q_frame, text="Dấu tách:", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE_ALT, fg=Theme.TEXT_SECONDARY).grid(row=1, column=0, sticky="w", pady=2)
        self.q_delim_var = tk.StringVar(value=".")
        ttk.Combobox(q_frame, textvariable=self.q_delim_var, values=[".", ":", ")", "-"], width=14).grid(row=1, column=1, padx=4, pady=2, sticky="ew")

        tk.Label(q_frame, text="Màu số:", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE_ALT, fg=Theme.TEXT_SECONDARY).grid(row=2, column=0, sticky="w", pady=2)
        self.q_color_var = tk.StringVar(value=self.COLOR_CHOICES[0])
        ttk.Combobox(q_frame, textvariable=self.q_color_var, values=self.COLOR_CHOICES, width=14).grid(row=2, column=1, padx=4, pady=2, sticky="ew")

        tk.Label(q_frame, text="Màu ABCD:", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE_ALT, fg=Theme.TEXT_SECONDARY).grid(row=3, column=0, sticky="w", pady=2)
        self.opt_color_var = tk.StringVar(value=self.COLOR_CHOICES[0])
        ttk.Combobox(q_frame, textvariable=self.opt_color_var, values=self.COLOR_CHOICES, width=14).grid(row=3, column=1, padx=4, pady=2, sticky="ew")

        tk.Label(q_frame, text="Màu [ins]:", font=Theme.FONT_SMALL, bg=Theme.BG_SURFACE_ALT, fg=Theme.TEXT_SECONDARY).grid(row=4, column=0, sticky="w", pady=2)
        self.ins_color_var = tk.StringVar(value=self.COLOR_CHOICES[0])
        ttk.Combobox(q_frame, textvariable=self.ins_color_var, values=self.COLOR_CHOICES, width=14).grid(row=4, column=1, padx=4, pady=2, sticky="ew")

        # ── Toggle Switches ──────────────────────────────────────
        self.pg_num_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            self,
            text="Đánh số trang (Trang X / Y)",
            variable=self.pg_num_var,
            font=Theme.FONT_BODY,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_PRIMARY,
            selectcolor=Theme.BG_SURFACE,
            activebackground=Theme.BG_SURFACE,
            highlightthickness=0,
            bd=0
        ).pack(anchor="w", pady=2)

        self.open_word_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            self,
            text="Mở file trong Word sau khi lưu",
            variable=self.open_word_var,
            font=Theme.FONT_BODY,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_PRIMARY,
            selectcolor=Theme.BG_SURFACE,
            activebackground=Theme.BG_SURFACE,
            highlightthickness=0,
            bd=0
        ).pack(anchor="w", pady=2)

        self.bg_mode_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            self,
            text="Chạy ngầm (Không giật chuột)",
            variable=self.bg_mode_var,
            font=Theme.FONT_BODY,
            bg=Theme.BG_SURFACE,
            fg=Theme.PRIMARY,
            selectcolor=Theme.BG_SURFACE,
            activebackground=Theme.BG_SURFACE,
            highlightthickness=0,
            bd=0
        ).pack(anchor="w", pady=2)

    def get_settings(self) -> Dict[str, Any]:
        """Returns clean dictionary of active settings."""
        return {
            "font_name": self.font_var.get(),
            "font_size": self.size_var.get(),
            "margin_top": self.m_top_var.get(),
            "margin_bottom": self.m_bottom_var.get(),
            "margin_left": self.m_left_var.get(),
            "margin_right": self.m_right_var.get(),
            "enable_page_numbers": self.pg_num_var.get(),
            "question_prefix": self.q_prefix_var.get(),
            "question_delimiter": self.q_delim_var.get(),
            "question_color": self.q_color_var.get(),
            "opt_color": self.opt_color_var.get(),
            "instruction_color": self.ins_color_var.get(),
            "ins_color": self.ins_color_var.get(),
            "open_word": self.open_word_var.get(),
            "background_mode": self.bg_mode_var.get(),
        }
