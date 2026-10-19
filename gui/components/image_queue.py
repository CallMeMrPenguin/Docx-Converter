"""
gui/components/image_queue.py - Flat Minimal Image Queue Component.
Manages user-selected image sequence for [PIC] tags, thumbnail preview,
and reordering with authentic zero-border Light UI styling.
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog
from typing import List, Optional, Callable
from PIL import Image, ImageTk

from gui.theme import Theme
from services.image_service import SUPPORTED_IMAGE_EXTENSIONS, ImageService
from core.utils import natural_sort_key


class ImageQueue(tk.Frame):
    """Component managing image files queue, live thumbnail, and ordering."""

    def __init__(self, parent, on_preview_requested: Optional[Callable[[int], None]] = None, on_match_requested: Optional[Callable[[], None]] = None, **kwargs):
        super().__init__(parent, bg=Theme.BG_SURFACE, padx=14, pady=10, **kwargs)
        self.on_preview_requested = on_preview_requested
        self.on_match_requested = on_match_requested
        self.user_image_paths: List[str] = []
        self._current_preview_photo: Optional[ImageTk.PhotoImage] = None

        # ── Title & Badge ────────────────────────────────────────
        header = tk.Frame(self, bg=Theme.BG_SURFACE)
        header.pack(fill="x", pady=(0, 6))

        tk.Label(
            header,
            text="Danh sách ảnh [PIC]",
            font=Theme.FONT_SECTION,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_PRIMARY
        ).pack(side="left")

        self.count_badge = tk.Label(
            header,
            text="0 ảnh",
            font=Theme.FONT_SMALL_BOLD,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.TEXT_SECONDARY,
            padx=6,
            pady=1
        )
        self.count_badge.pack(side="right")

        # ── Image Listbox Container ──────────────────────────────
        list_container = tk.Frame(self, bg=Theme.BG_SURFACE_ALT, padx=2, pady=2)
        list_container.pack(fill="x", pady=(0, 6))

        self.img_listbox = tk.Listbox(
            list_container,
            height=4,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_PRIMARY,
            font=Theme.FONT_BODY,
            selectbackground=Theme.PRIMARY_LIGHT,
            selectforeground=Theme.PRIMARY,
            activestyle="none",
            highlightthickness=0,
            bd=0,
            relief="flat"
        )
        self.img_listbox.pack(side="left", fill="both", expand=True)

        list_scroll = ttk.Scrollbar(list_container, orient="vertical", command=self.img_listbox.yview)
        list_scroll.pack(side="right", fill="y")
        self.img_listbox.config(yscrollcommand=list_scroll.set)

        self.img_listbox.bind("<<ListboxSelect>>", self._on_selection_change)
        self.img_listbox.bind("<Double-Button-1>", lambda e: self._on_double_click())
        self.img_listbox.bind("<KeyRelease-Up>", self._on_selection_change)
        self.img_listbox.bind("<KeyRelease-Down>", self._on_selection_change)

        # ── Action Buttons Row 1: Add, Sort, Clear ──────────────
        btn_row1 = tk.Frame(self, bg=Theme.BG_SURFACE)
        btn_row1.pack(fill="x", pady=(0, 4))

        btn_add = tk.Button(
            btn_row1,
            text="+ Thêm ảnh",
            command=self.add_images,
            bg=Theme.PRIMARY,
            fg=Theme.TEXT_ON_PRIMARY,
            activebackground=Theme.PRIMARY_HOVER,
            activeforeground=Theme.TEXT_ON_PRIMARY,
            font=Theme.FONT_SMALL_BOLD,
            relief="flat",
            bd=0,
            pady=4,
            cursor="hand2"
        )
        btn_add.pack(side="left", fill="x", expand=True, padx=(0, 2))

        btn_sort = tk.Button(
            btn_row1,
            text="A-Z",
            command=self.sort_images,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.TEXT_PRIMARY,
            activebackground=Theme.BG_HOVER,
            font=Theme.FONT_SMALL_BOLD,
            relief="flat",
            bd=0,
            pady=4,
            cursor="hand2"
        )
        btn_sort.pack(side="left", padx=2)

        btn_clear = tk.Button(
            btn_row1,
            text="Xóa hết",
            command=self.clear_images,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.DANGER,
            activebackground=Theme.BG_HOVER,
            font=Theme.FONT_SMALL_BOLD,
            relief="flat",
            bd=0,
            pady=4,
            cursor="hand2"
        )
        btn_clear.pack(side="right", padx=(2, 0))

        # ── Action Buttons Row 2: Up, Down, Remove, View ─────────
        btn_row2 = tk.Frame(self, bg=Theme.BG_SURFACE)
        btn_row2.pack(fill="x", pady=(0, 6))

        btn_up = tk.Button(
            btn_row2,
            text="▲ Lên",
            command=self.move_image_up,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.TEXT_PRIMARY,
            activebackground=Theme.BG_HOVER,
            font=Theme.FONT_SMALL,
            relief="flat",
            bd=0,
            pady=3,
            cursor="hand2"
        )
        btn_up.pack(side="left", fill="x", expand=True, padx=(0, 2))

        btn_down = tk.Button(
            btn_row2,
            text="▼ Xuống",
            command=self.move_image_down,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.TEXT_PRIMARY,
            activebackground=Theme.BG_HOVER,
            font=Theme.FONT_SMALL,
            relief="flat",
            bd=0,
            pady=3,
            cursor="hand2"
        )
        btn_down.pack(side="left", fill="x", expand=True, padx=2)

        btn_del = tk.Button(
            btn_row2,
            text="Xóa",
            command=self.remove_selected_image,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.DANGER,
            activebackground=Theme.BG_HOVER,
            font=Theme.FONT_SMALL,
            relief="flat",
            bd=0,
            pady=3,
            cursor="hand2"
        )
        btn_del.pack(side="left", fill="x", expand=True, padx=2)

        btn_view = tk.Button(
            btn_row2,
            text="Xem lớn",
            command=self._on_double_click,
            bg=Theme.INFO,
            fg=Theme.TEXT_ON_PRIMARY,
            activebackground=Theme.INFO_HOVER,
            activeforeground=Theme.TEXT_ON_PRIMARY,
            font=Theme.FONT_SMALL_BOLD,
            relief="flat",
            bd=0,
            pady=3,
            cursor="hand2"
        )
        btn_view.pack(side="right", fill="x", expand=True, padx=(2, 0))

        # ── Exercise Match Preview Button ────────────────────────
        btn_match = tk.Button(
            self,
            text="👁️ Đối chiếu câu hỏi & ảnh",
            command=self._trigger_match,
            bg=Theme.PURPLE,
            fg=Theme.TEXT_ON_PRIMARY,
            activebackground=Theme.PURPLE_HOVER,
            activeforeground=Theme.TEXT_ON_PRIMARY,
            font=Theme.FONT_SMALL_BOLD,
            relief="flat",
            bd=0,
            pady=4,
            cursor="hand2"
        )
        btn_match.pack(fill="x", pady=(0, 6))

        # ── Live Thumbnail Preview Box ───────────────────────────
        self.preview_card = tk.Frame(self, bg=Theme.BG_SURFACE_ALT, pady=6, padx=6)
        self.preview_card.pack(fill="x")

        self.preview_canvas = tk.Canvas(self.preview_card, height=100, bg=Theme.BG_SURFACE_ALT, highlightthickness=0, bd=0, cursor="hand2")
        self.preview_canvas.pack(fill="x")
        self.preview_canvas.bind("<Button-1>", lambda e: self._on_double_click())

        self.preview_info_lbl = tk.Label(
            self.preview_card,
            text="Chưa chọn ảnh nào",
            font=Theme.FONT_SMALL,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.TEXT_MUTED
        )
        self.preview_info_lbl.pack(fill="x", pady=(4, 0))

    def add_images(self):
        """Allows user to select multiple images and appends them cleanly."""
        types = [
            ("Hình ảnh", "*.png;*.jpg;*.jpeg;*.webp;*.bmp;*.gif;*.tif;*.tiff;*.heic;*.heif;*.avif"),
            ("Tất cả tệp", "*.*")
        ]
        files = filedialog.askopenfilenames(title="Chọn ảnh cho tài liệu", filetypes=types)
        if files:
            for f in files:
                abs_p = os.path.abspath(f)
                if abs_p not in self.user_image_paths:
                    self.user_image_paths.append(abs_p)
            self.update_image_listbox(selected_index=len(self.user_image_paths) - 1)

    def sort_images(self):
        """Sorts images alphabetically by filename using natural sort order."""
        if not self.user_image_paths:
            return
        self.user_image_paths.sort(key=lambda p: natural_sort_key(os.path.basename(p)))
        self.update_image_listbox(selected_index=0)

    def clear_images(self):
        """Clears all images from the queue."""
        self.user_image_paths.clear()
        self.update_image_listbox()

    def move_image_up(self):
        sel = self.img_listbox.curselection()
        if not sel or sel[0] <= 0:
            return
        idx = sel[0]
        self.user_image_paths[idx - 1], self.user_image_paths[idx] = self.user_image_paths[idx], self.user_image_paths[idx - 1]
        self.update_image_listbox(selected_index=idx - 1)

    def move_image_down(self):
        sel = self.img_listbox.curselection()
        if not sel or sel[0] >= len(self.user_image_paths) - 1:
            return
        idx = sel[0]
        self.user_image_paths[idx + 1], self.user_image_paths[idx] = self.user_image_paths[idx], self.user_image_paths[idx + 1]
        self.update_image_listbox(selected_index=idx + 1)

    def remove_selected_image(self):
        sel = self.img_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        self.user_image_paths.pop(idx)
        new_idx = min(idx, len(self.user_image_paths) - 1) if self.user_image_paths else None
        self.update_image_listbox(selected_index=new_idx)

    def update_image_listbox(self, selected_index: Optional[int] = None):
        """Refreshes listbox contents and updates selection."""
        self.img_listbox.delete(0, tk.END)
        for idx, p in enumerate(self.user_image_paths, 1):
            self.img_listbox.insert(tk.END, f"{idx}. {os.path.basename(p)}")

        self.count_badge.config(text=f"{len(self.user_image_paths)} ảnh")

        if self.user_image_paths:
            if selected_index is not None and 0 <= selected_index < len(self.user_image_paths):
                self.img_listbox.selection_set(selected_index)
                self.img_listbox.see(selected_index)
            elif not self.img_listbox.curselection():
                self.img_listbox.selection_set(0)
            self._render_thumbnail()
        else:
            self.preview_canvas.delete("all")
            self.preview_info_lbl.config(text="Chưa chọn ảnh nào")

    def _on_selection_change(self, event=None):
        self._render_thumbnail()

    def _on_double_click(self):
        sel = self.img_listbox.curselection()
        if sel and self.on_preview_requested:
            self.on_preview_requested(sel[0])

    def _trigger_match(self):
        if self.on_match_requested:
            self.on_match_requested()

    def _render_thumbnail(self):
        """Renders scaled thumbnail in canvas for currently selected image."""
        sel = self.img_listbox.curselection()
        if not sel or sel[0] >= len(self.user_image_paths):
            self.preview_canvas.delete("all")
            self.preview_info_lbl.config(text="Chưa chọn ảnh nào")
            return

        fpath = self.user_image_paths[sel[0]]
        if not os.path.exists(fpath):
            self.preview_canvas.delete("all")
            self.preview_info_lbl.config(text="[Tệp không tồn tại]")
            return

        try:
            pil_thumb = ImageService.create_thumbnail(fpath, max_size=(180, 95))
            if pil_thumb:
                self._current_preview_photo = ImageTk.PhotoImage(pil_thumb)
                self.preview_canvas.delete("all")
                c_w = self.preview_canvas.winfo_width() or 180
                c_h = self.preview_canvas.winfo_height() or 95
                self.preview_canvas.create_image(c_w // 2, c_h // 2, image=self._current_preview_photo, anchor="center")
                sz_kb = os.path.getsize(fpath) / 1024.0
                self.preview_info_lbl.config(text=f"{os.path.basename(fpath)} ({sz_kb:.0f} KB)")
        except Exception as e:
            self.preview_canvas.delete("all")
            self.preview_info_lbl.config(text=f"Lỗi: {e}")

    def get_image_paths(self) -> List[str]:
        return list(self.user_image_paths)

    def set_image_paths(self, paths: List[str]):
        self.user_image_paths = list(paths)
        self.update_image_listbox(selected_index=0 if paths else None)
