"""
gui/components/search_bar.py - Flat Minimal Search & Replace Toolbar Component.
Provides real-time debounced text search, replace all, and navigation without component borders.
"""

import tkinter as tk
from typing import Optional, Callable
from gui.theme import Theme


class SearchBar(tk.Frame):
    """Clean, flat toolbar for finding and replacing text in the ULN editor."""

    def __init__(
        self,
        parent,
        text_widget: tk.Text,
        on_close: Optional[Callable[[], None]] = None,
        **kwargs
    ):
        super().__init__(parent, bg=Theme.BG_SURFACE_ALT, padx=12, pady=6, **kwargs)
        self.text_widget = text_widget
        self.on_close_callback = on_close

        self.matches = []
        self.current_match_idx = -1
        self._debounce_job = None
        self.replace_mode = False

        # ── Find Row ─────────────────────────────────────────────
        find_row = tk.Frame(self, bg=Theme.BG_SURFACE_ALT)
        find_row.pack(fill="x", pady=(0, 2))

        tk.Label(
            find_row,
            text="Tìm kiếm:",
            font=Theme.FONT_SMALL_BOLD,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.TEXT_SECONDARY
        ).pack(side="left", padx=(0, 6))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self._schedule_search())

        self.search_entry = tk.Entry(
            find_row,
            textvariable=self.search_var,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_PRIMARY,
            insertbackground=Theme.PRIMARY,
            font=Theme.FONT_CODE,
            relief="flat",
            bd=0,
            highlightthickness=0,
            width=24
        )
        self.search_entry.pack(side="left", padx=(0, 6), ipady=3)
        self.search_entry.bind("<Return>", lambda e: self.find_next())
        self.search_entry.bind("<Shift-Return>", lambda e: self.find_prev())

        self.match_count_lbl = tk.Label(
            find_row,
            text="0 kết quả",
            font=Theme.FONT_SMALL,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.TEXT_MUTED,
            width=10,
            anchor="w"
        )
        self.match_count_lbl.pack(side="left", padx=(0, 6))

        btn_prev = tk.Button(
            find_row,
            text="▲",
            command=self.find_prev,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_PRIMARY,
            activebackground=Theme.BG_HOVER,
            font=Theme.FONT_SMALL_BOLD,
            relief="flat",
            bd=0,
            padx=6,
            pady=2,
            cursor="hand2"
        )
        btn_prev.pack(side="left", padx=1)

        btn_next = tk.Button(
            find_row,
            text="▼",
            command=self.find_next,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_PRIMARY,
            activebackground=Theme.BG_HOVER,
            font=Theme.FONT_SMALL_BOLD,
            relief="flat",
            bd=0,
            padx=6,
            pady=2,
            cursor="hand2"
        )
        btn_next.pack(side="left", padx=1)

        self.match_case_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            find_row,
            text="Aa Hoa/thường",
            variable=self.match_case_var,
            command=self._do_search,
            font=Theme.FONT_SMALL,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.TEXT_SECONDARY,
            selectcolor=Theme.BG_SURFACE,
            activebackground=Theme.BG_SURFACE_ALT,
            bd=0,
            highlightthickness=0
        ).pack(side="left", padx=6)

        self.btn_replace_toggle = tk.Button(
            find_row,
            text="Thay thế...",
            command=self.toggle_replace_mode,
            bg=Theme.BG_SURFACE,
            fg=Theme.WARNING,
            activebackground=Theme.BG_HOVER,
            font=Theme.FONT_SMALL_BOLD,
            relief="flat",
            bd=0,
            padx=8,
            pady=2,
            cursor="hand2"
        )
        self.btn_replace_toggle.pack(side="left", padx=4)

        btn_close = tk.Button(
            find_row,
            text="✕",
            command=self.close,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.TEXT_MUTED,
            activebackground=Theme.BG_SURFACE_ALT,
            activeforeground=Theme.DANGER,
            font=Theme.FONT_BODY_BOLD,
            relief="flat",
            bd=0,
            padx=6,
            cursor="hand2"
        )
        btn_close.pack(side="right")

        # ── Replace Row (Conditional) ────────────────────────────
        self.replace_row = tk.Frame(self, bg=Theme.BG_SURFACE_ALT)

        tk.Label(
            self.replace_row,
            text="Thay bằng:",
            font=Theme.FONT_SMALL_BOLD,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.TEXT_SECONDARY
        ).pack(side="left", padx=(0, 6))

        self.replace_var = tk.StringVar()
        self.replace_entry = tk.Entry(
            self.replace_row,
            textvariable=self.replace_var,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_PRIMARY,
            insertbackground=Theme.PRIMARY,
            font=Theme.FONT_CODE,
            relief="flat",
            bd=0,
            highlightthickness=0,
            width=24
        )
        self.replace_entry.pack(side="left", padx=(0, 6), ipady=3)
        self.replace_entry.bind("<Return>", lambda e: self.replace_current())

        btn_rep_one = tk.Button(
            self.replace_row,
            text="Thay thế",
            command=self.replace_current,
            bg=Theme.WARNING,
            fg=Theme.TEXT_ON_PRIMARY,
            activebackground=Theme.WARNING_HOVER,
            activeforeground=Theme.TEXT_ON_PRIMARY,
            font=Theme.FONT_SMALL_BOLD,
            relief="flat",
            bd=0,
            padx=8,
            pady=2,
            cursor="hand2"
        )
        btn_rep_one.pack(side="left", padx=2)

        btn_rep_all = tk.Button(
            self.replace_row,
            text="Thay tất cả",
            command=self.replace_all,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_PRIMARY,
            activebackground=Theme.BG_HOVER,
            font=Theme.FONT_SMALL_BOLD,
            relief="flat",
            bd=0,
            padx=8,
            pady=2,
            cursor="hand2"
        )
        btn_rep_all.pack(side="left", padx=2)

    def _schedule_search(self):
        """Debounces search query by 200ms for smooth instantaneous responsiveness."""
        if self._debounce_job:
            self.after_cancel(self._debounce_job)
        self._debounce_job = self.after(200, self._do_search)

    def _do_search(self):
        query = self.search_var.get()
        self.text_widget.tag_remove("search_highlight", "1.0", tk.END)
        self.text_widget.tag_remove("search_current", "1.0", tk.END)
        self.matches.clear()
        self.current_match_idx = -1

        if not query:
            self.match_count_lbl.config(text="0 kết quả")
            return

        nocase = not self.match_case_var.get()
        start = "1.0"
        while True:
            pos = self.text_widget.search(query, start, stopindex=tk.END, nocase=nocase)
            if not pos:
                break
            end = f"{pos}+{len(query)}c"
            self.matches.append((pos, end))
            start = end

        # Configure search highlight tags
        self.text_widget.tag_config("search_highlight", background="#fef08a", foreground="#000000")
        self.text_widget.tag_config("search_current", background="#f97316", foreground="#ffffff")

        for pos, end in self.matches:
            self.text_widget.tag_add("search_highlight", pos, end)

        if self.matches:
            self.current_match_idx = 0
            self._highlight_current()
            self.match_count_lbl.config(text=f"1 / {len(self.matches)}")
        else:
            self.match_count_lbl.config(text="0 kết quả")

    def _highlight_current(self):
        if not self.matches or self.current_match_idx < 0:
            return
        self.text_widget.tag_remove("search_current", "1.0", tk.END)
        pos, end = self.matches[self.current_match_idx]
        self.text_widget.tag_add("search_current", pos, end)
        self.text_widget.see(pos)
        self.match_count_lbl.config(text=f"{self.current_match_idx + 1} / {len(self.matches)}")

    def find_next(self):
        if not self.matches:
            return
        self.current_match_idx = (self.current_match_idx + 1) % len(self.matches)
        self._highlight_current()

    def find_prev(self):
        if not self.matches:
            return
        self.current_match_idx = (self.current_match_idx - 1) % len(self.matches)
        self._highlight_current()

    def toggle_replace_mode(self):
        self.replace_mode = not self.replace_mode
        if self.replace_mode:
            self.replace_row.pack(fill="x", pady=(4, 0))
            self.btn_replace_toggle.config(bg=Theme.WARNING, fg=Theme.TEXT_ON_PRIMARY)
            self.replace_entry.focus_set()
        else:
            self.replace_row.pack_forget()
            self.btn_replace_toggle.config(bg=Theme.BG_SURFACE, fg=Theme.WARNING)

    def replace_current(self):
        if not self.matches or self.current_match_idx < 0:
            return
        rep_val = self.replace_var.get()
        pos, end = self.matches[self.current_match_idx]
        self.text_widget.delete(pos, end)
        self.text_widget.insert(pos, rep_val)
        self._do_search()

    def replace_all(self):
        if not self.matches:
            return
        query = self.search_var.get()
        rep_val = self.replace_var.get()
        content = self.text_widget.get("1.0", tk.END)
        if not self.match_case_var.get():
            import re
            new_content = re.sub(re.escape(query), rep_val, content, flags=re.IGNORECASE)
        else:
            new_content = content.replace(query, rep_val)
        self.text_widget.delete("1.0", tk.END)
        self.text_widget.insert("1.0", new_content)
        self._do_search()

    def close(self):
        self.text_widget.tag_remove("search_highlight", "1.0", tk.END)
        self.text_widget.tag_remove("search_current", "1.0", tk.END)
        self.pack_forget()
        if self.on_close_callback:
            self.on_close_callback()
