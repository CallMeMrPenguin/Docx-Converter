"""
gui/app.py - Main Desktop Application Coordinator.
Integrates Flat Minimal Light UI theme, modular components, non-blocking background
compilation thread, and safe cancellation tokens.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from typing import Optional

# Enable per-monitor DPI awareness on Windows before Tk initialization
try:
    import ctypes
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

from gui.theme import Theme, apply_light_theme
from gui.components.settings_panel import SettingsPanel
from gui.components.image_queue import ImageQueue
from gui.components.search_bar import SearchBar
from gui.components.status_bar import StatusBar

from gui.dialogs.prompt_editor import open_prompt_editor_dialog
from gui.dialogs.update_modal import show_update_modal_dialog
from gui.dialogs.docx_picker import open_docx_picker_dialog
from gui.dialogs.image_preview import open_image_preview_dialog, open_exercise_pic_match_dialog

from core.compiler import ULNCompiler, embed_raw_uln_docx
import updater


class ULNFormatterApp:
    """Main Application window coordinating editor, sidebar, and background compiler."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.version = updater.get_current_version()
        self.root.title(f"ULN to DOCX Converter v{self.version}")
        self.root.geometry("1260x860")
        self.root.minsize(1050, 700)
        self.root.configure(bg=Theme.BG_CANVAS)

        self.last_output_dir = os.path.expanduser("~\\Documents")
        if not os.path.exists(self.last_output_dir):
            self.last_output_dir = os.getcwd()

        self._compiling_thread: Optional[threading.Thread] = None
        self._cancel_event = threading.Event()

        # Apply Light UI theme tokens
        apply_light_theme(self.root)

        # ── 1. Top Header Bar ────────────────────────────────────────
        header = tk.Frame(self.root, bg=Theme.BG_SURFACE, height=54, padx=20, pady=10)
        header.pack(fill="x", side="top")

        # Left branding
        title_box = tk.Frame(header, bg=Theme.BG_SURFACE)
        title_box.pack(side="left", fill="y")

        tk.Label(
            title_box,
            text="ULN to DOCX Converter",
            font=Theme.FONT_TITLE,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_PRIMARY
        ).pack(side="left")

        tk.Label(
            title_box,
            text=f"v{self.version}",
            font=Theme.FONT_SMALL_BOLD,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.PRIMARY,
            padx=6,
            pady=2
        ).pack(side="left", padx=8)

        tk.Label(
            title_box,
            text="Universal Layout Notation",
            font=Theme.FONT_SUBTITLE,
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_MUTED
        ).pack(side="left", padx=4)

        # Right actions (Prompt & Update)
        actions_box = tk.Frame(header, bg=Theme.BG_SURFACE)
        actions_box.pack(side="right", fill="y")

        tk.Button(
            actions_box,
            text="📜 Quy tắc & Prompt AI",
            command=self.open_prompt_editor,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.TEXT_PRIMARY,
            activebackground=Theme.BG_HOVER,
            font=Theme.FONT_BODY_BOLD,
            relief="flat",
            bd=0,
            padx=12,
            pady=4,
            cursor="hand2"
        ).pack(side="left", padx=(0, 6))

        tk.Button(
            actions_box,
            text="🔄 Kiểm tra Cập nhật",
            command=self.manual_check_updates,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.PRIMARY,
            activebackground=Theme.BG_HOVER,
            font=Theme.FONT_BODY_BOLD,
            relief="flat",
            bd=0,
            padx=10,
            pady=4,
            cursor="hand2"
        ).pack(side="left")

        # ── 2. Main Workspace (Sidebar + Editor Area) ────────────────
        main_workspace = tk.Frame(self.root, bg=Theme.BG_CANVAS, padx=16, pady=12)
        main_workspace.pack(fill="both", expand=True)

        # Left Sidebar (Scrollable Container for Settings + Image Queue)
        sidebar_outer = tk.Frame(main_workspace, bg=Theme.BG_SURFACE, width=320)
        sidebar_outer.pack(side="left", fill="y", padx=(0, 12))
        sidebar_outer.pack_propagate(False)

        side_canvas = tk.Canvas(sidebar_outer, bg=Theme.BG_SURFACE, highlightthickness=0, bd=0)
        side_scroll = ttk.Scrollbar(sidebar_outer, orient="vertical", command=side_canvas.yview)
        self.sidebar_frame = tk.Frame(side_canvas, bg=Theme.BG_SURFACE)

        self.sidebar_frame.bind(
            "<Configure>",
            lambda e: side_canvas.configure(scrollregion=side_canvas.bbox("all"))
        )
        side_window = side_canvas.create_window((0, 0), window=self.sidebar_frame, anchor="nw")

        def _on_side_canvas_configure(e):
            side_canvas.itemconfig(side_window, width=e.width)
        side_canvas.bind("<Configure>", _on_side_canvas_configure)
        side_canvas.configure(yscrollcommand=side_scroll.set)

        side_canvas.pack(side="left", fill="both", expand=True)
        side_scroll.pack(side="right", fill="y")

        # Mount Settings Panel
        self.settings_panel = SettingsPanel(self.sidebar_frame)
        self.settings_panel.pack(fill="x")

        # Mount Image Queue
        self.image_queue = ImageQueue(
            self.sidebar_frame,
            on_preview_requested=self.open_image_preview,
            on_match_requested=self.open_pic_match_dialog
        )
        self.image_queue.pack(fill="x", pady=(10, 10))

        # Right Area: Editor Container Card
        editor_card = tk.Frame(main_workspace, bg=Theme.BG_SURFACE, padx=12, pady=10)
        editor_card.pack(side="right", fill="both", expand=True)

        # ── 3. Toolbar inside Editor Card ────────────────────────────
        toolbar = tk.Frame(editor_card, bg=Theme.BG_SURFACE, pady=4)
        toolbar.pack(fill="x", side="top")

        # Compile Button (Prominently styled in Emerald Green)
        self.btn_compile = tk.Button(
            toolbar,
            text="🚀 XUẤT FILE DOCX",
            command=self.compile_docx,
            bg=Theme.SUCCESS,
            fg=Theme.TEXT_ON_PRIMARY,
            activebackground=Theme.SUCCESS_HOVER,
            activeforeground=Theme.TEXT_ON_PRIMARY,
            font=Theme.FONT_BODY_BOLD,
            relief="flat",
            bd=0,
            padx=16,
            pady=5,
            cursor="hand2"
        )
        self.btn_compile.pack(side="right", padx=(8, 0))

        tk.Button(
            toolbar,
            text="🗑️ Xóa",
            command=self.clear_text,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.DANGER,
            activebackground=Theme.BG_HOVER,
            font=Theme.FONT_SMALL_BOLD,
            relief="flat",
            bd=0,
            padx=8,
            pady=4,
            cursor="hand2"
        ).pack(side="right", padx=2)

        self.btn_search = tk.Button(
            toolbar,
            text="🔍 Tìm (Ctrl+F)",
            command=self.toggle_search_bar,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.PRIMARY,
            activebackground=Theme.BG_HOVER,
            font=Theme.FONT_SMALL_BOLD,
            relief="flat",
            bd=0,
            padx=8,
            pady=4,
            cursor="hand2"
        )
        self.btn_search.pack(side="right", padx=2)

        tk.Button(
            toolbar,
            text="📥 Nhập DOCX",
            command=self.open_docx_picker,
            bg=Theme.INFO,
            fg=Theme.TEXT_ON_PRIMARY,
            activebackground=Theme.INFO_HOVER,
            activeforeground=Theme.TEXT_ON_PRIMARY,
            font=Theme.FONT_SMALL_BOLD,
            relief="flat",
            bd=0,
            padx=9,
            pady=4,
            cursor="hand2"
        ).pack(side="left", padx=(0, 2))

        tk.Button(
            toolbar,
            text="💾 Nhúng Raw",
            command=self.embed_raw_to_docx_file,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.TEXT_PRIMARY,
            activebackground=Theme.BG_HOVER,
            font=Theme.FONT_SMALL_BOLD,
            relief="flat",
            bd=0,
            padx=8,
            pady=4,
            cursor="hand2"
        ).pack(side="left", padx=2)

        tk.Button(
            toolbar,
            text="📁 Mở tệp .txt",
            command=self.import_file,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.TEXT_PRIMARY,
            activebackground=Theme.BG_HOVER,
            font=Theme.FONT_SMALL_BOLD,
            relief="flat",
            bd=0,
            padx=8,
            pady=4,
            cursor="hand2"
        ).pack(side="left", padx=2)

        tk.Button(
            toolbar,
            text="📄 Mẫu ULN",
            command=self.load_sample,
            bg=Theme.BG_SURFACE_ALT,
            fg=Theme.TEXT_SECONDARY,
            activebackground=Theme.BG_HOVER,
            font=Theme.FONT_SMALL_BOLD,
            relief="flat",
            bd=0,
            padx=8,
            pady=4,
            cursor="hand2"
        ).pack(side="left", padx=2)

        # Quick Heading Shortcuts (H1-H6)
        hdr_frame = tk.Frame(toolbar, bg=Theme.BG_SURFACE)
        hdr_frame.pack(side="left", padx=(8, 0))
        tk.Label(hdr_frame, text="H:", font=Theme.FONT_SMALL_BOLD, bg=Theme.BG_SURFACE, fg=Theme.TEXT_MUTED).pack(side="left", padx=(0, 2))
        for lvl in range(1, 7):
            tk.Button(
                hdr_frame,
                text=f"H{lvl}",
                command=lambda l=lvl: self.apply_heading_shortcut(l),
                bg=Theme.BG_SURFACE_ALT,
                fg=Theme.TEXT_SECONDARY,
                activebackground=Theme.PRIMARY_LIGHT,
                activeforeground=Theme.PRIMARY,
                font=Theme.FONT_SMALL_BOLD,
                relief="flat",
                bd=0,
                padx=4,
                pady=2,
                cursor="hand2"
            ).pack(side="left", padx=1)

        # ── 4. Search Bar (Hidden by default) ────────────────────────
        self.search_bar = SearchBar(editor_card, text_widget=None, on_close=self._on_search_closed)

        # ── 5. Main Text Editor ──────────────────────────────────────
        editor_container = tk.Frame(editor_card, bg=Theme.BG_SURFACE_ALT, padx=2, pady=2)
        editor_container.pack(fill="both", expand=True, pady=(6, 4))

        txt_scroll = ttk.Scrollbar(editor_container)
        txt_scroll.pack(side="right", fill="y")

        self.text_editor = tk.Text(
            editor_container,
            wrap="none",
            bg=Theme.BG_SURFACE,
            fg=Theme.TEXT_PRIMARY,
            insertbackground=Theme.PRIMARY,
            selectbackground=Theme.PRIMARY_LIGHT,
            selectforeground=Theme.PRIMARY,
            font=Theme.FONT_CODE_EDITOR,
            yscrollcommand=txt_scroll.set,
            padx=12,
            pady=10,
            bd=0,
            highlightthickness=0,
            relief="flat",
            undo=True
        )
        self.text_editor.pack(fill="both", expand=True)
        txt_scroll.config(command=self.text_editor.yview)

        # Link search bar with text editor
        self.search_bar.text_widget = self.text_editor

        # Live stats tracking
        self.text_editor.bind("<KeyRelease>", lambda e: self.update_status_metrics())
        self.text_editor.bind("<<Modified>>", lambda e: self._on_text_modified())

        # Keyboard shortcuts
        self.root.bind("<Control-f>", lambda e: self.show_search_bar(replace=False))
        self.root.bind("<Control-h>", lambda e: self.show_search_bar(replace=True))

        for lvl in range(1, 7):
            self.root.bind(f"<Alt-Key-{lvl}>", lambda e, l=lvl: self.apply_heading_shortcut(l))

        # ── 6. Bottom Status Bar ─────────────────────────────────────
        self.status_bar = StatusBar(self.root, on_cancel=self.cancel_compilation)
        self.status_bar.pack(fill="x", side="bottom")

        # Load initial sample or welcome text
        self.load_sample()
        self.update_status_metrics()

    @property
    def user_image_paths(self):
        return self.image_queue.get_image_paths()

    def update_status_metrics(self):
        content = self.text_editor.get("1.0", tk.END).rstrip("\n")
        self.status_bar.update_stats(content)

    def _on_text_modified(self):
        if self.text_editor.edit_modified():
            self.update_status_metrics()
            self.text_editor.edit_modified(False)

    def show_search_bar(self, replace: bool = False):
        self.search_bar.pack(fill="x", side="top", pady=(0, 4))
        if replace and not self.search_bar.replace_mode:
            self.search_bar.toggle_replace_mode()
        self.search_bar.search_entry.focus_set()
        return "break"

    def toggle_search_bar(self):
        if self.search_bar.winfo_ismapped():
            self.search_bar.close()
        else:
            self.show_search_bar(replace=False)

    def _on_search_closed(self):
        self.text_editor.focus_set()

    def apply_heading_shortcut(self, level: int):
        try:
            cur_line = self.text_editor.get("insert linestart", "insert lineend")
            clean = cur_line.lstrip()
            for k in range(1, 7):
                if clean.startswith(f"[H{k}]"):
                    clean = clean[5:].lstrip()
                    break
            new_line = f"[H{level}] {clean}"
            self.text_editor.delete("insert linestart", "insert lineend")
            self.text_editor.insert("insert linestart", new_line)
        except Exception:
            pass

    def clear_text(self):
        self.text_editor.delete("1.0", tk.END)
        self.update_status_metrics()

    def load_sample(self):
        sample_path = os.path.join(updater.get_bundle_dir(), "uln_test.txt")
        if not os.path.exists(sample_path):
            sample_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "uln_test.txt")

        content = ""
        if os.path.exists(sample_path):
            try:
                with open(sample_path, "r", encoding="utf-8") as f:
                    content = f.read()
            except Exception:
                pass

        if not content:
            content = (
                "[H1] ĐỀ KIỂM TRA MÔN TIẾNG ANH\n"
                "[P0] Mark the letter A, B, C, or D to indicate the word whose underlined part differs:\n"
                "[P1] 1. A. [c]{u}at \t B. [c]{u}ity \t C. [c]{u}ar \t D. [c]{u}up\n"
                "[P1] 2. A. teach[er]{u} \t B. fath[er]{u} \t C. t[er]{u}m \t D. moth[er]{u}\n"
            )

        self.text_editor.delete("1.0", tk.END)
        self.text_editor.insert("1.0", content)
        self.update_status_metrics()

    def import_file(self):
        fpath = filedialog.askopenfilename(
            title="Chọn tệp văn bản ULN",
            filetypes=[("Text Files", "*.txt"), ("ULN Files", "*.uln"), ("All Files", "*.*")]
        )
        if fpath:
            try:
                with open(fpath, "r", encoding="utf-8") as f:
                    content = f.read()
                self.text_editor.delete("1.0", tk.END)
                self.text_editor.insert("1.0", content)
                self.update_status_metrics()
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể đọc file: {e}")

    def open_prompt_editor(self):
        open_prompt_editor_dialog(self.root)

    def manual_check_updates(self):
        self.status_bar.set_message("Đang kiểm tra bản cập nhật...")
        def _check():
            info = updater.check_for_updates()
            def _show():
                if info.get("has_update"):
                    show_update_modal_dialog(self.root, info)
                elif info.get("success"):
                    messagebox.showinfo("Cập nhật", f"Bạn đang sử dụng phiên bản mới nhất (v{info['current_version']}).")
                else:
                    messagebox.showwarning("Kiểm tra cập nhật", f"Không thể kiểm tra cập nhật:\n{info.get('error', 'Lỗi mạng')}")
                self.status_bar.set_message("Sẵn sàng")
            self.root.after(0, _show)
        threading.Thread(target=_check, daemon=True).start()

    def open_docx_picker(self):
        open_docx_picker_dialog(self)

    def open_image_preview(self, initial_index: int = 0):
        open_image_preview_dialog(self, initial_index=initial_index)

    def open_pic_match_dialog(self):
        open_exercise_pic_match_dialog(self)

    def embed_raw_to_docx_file(self):
        uln_text = self.text_editor.get("1.0", tk.END).strip()
        if not uln_text:
            messagebox.showwarning("Thông báo", "Trình soạn thảo đang trống.")
            return

        fpath = filedialog.askopenfilename(
            title="Chọn file DOCX để nhúng mã nguồn ULN",
            initialdir=self.last_output_dir,
            filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")]
        )
        if fpath and os.path.exists(fpath):
            ok = embed_raw_uln_docx(fpath, uln_text)
            if ok:
                messagebox.showinfo("Thành công", f"Đã nhúng thành công mã Raw ULN vào file:\n{os.path.basename(fpath)}")
            else:
                messagebox.showerror("Lỗi", "Không thể nhúng mã ULN. Vui lòng đảm bảo file đã được đóng trong Word.")

    def cancel_compilation(self):
        """Signals active background compilation thread to stop gracefully."""
        self._cancel_event.set()
        self.status_bar.set_message("Đang yêu cầu hủy tác vụ...")

    def compile_docx(self):
        """Starts non-blocking background compilation thread with live progress."""
        if self._compiling_thread and self._compiling_thread.is_alive():
            messagebox.showwarning("Đang xử lý", "Một tác vụ biên dịch DOCX đang chạy.")
            return

        uln_text = self.text_editor.get("1.0", tk.END).strip()
        if not uln_text:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập mã ULN trước khi biên dịch.")
            return

        ts_filename = f"uln_document_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        out_path = filedialog.asksaveasfilename(
            title="Lưu file DOCX đã định dạng",
            initialdir=self.last_output_dir,
            initialfile=ts_filename,
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx"), ("All Files", "*.*")]
        )
        if not out_path:
            return

        self.last_output_dir = os.path.dirname(os.path.abspath(out_path))

        # Build settings dict
        settings = self.settings_panel.get_settings()
        settings["user_images"] = self.image_queue.get_image_paths()
        self._cancel_event.clear()
        settings["cancel_event"] = self._cancel_event

        self.btn_compile.config(state="disabled", text="⏳ ĐANG BIÊN DỊCH...")
        self.status_bar.show_progress(0.05, "Đang khởi tạo Microsoft Word...")

        def _worker():
            err_msg = None
            is_cancelled = False
            compiled_path = None

            def _progress(msg, pct):
                self.root.after(0, lambda: self.status_bar.show_progress(pct, msg))

            try:
                compiler = ULNCompiler(settings)
                compiled_path = compiler.compile(
                    uln_text=uln_text,
                    output_filepath=out_path,
                    keep_open=settings.get("open_word", True),
                    background_mode=settings.get("background_mode", False),
                    cancel_check=lambda: self._cancel_event.is_set(),
                    progress_callback=_progress
                )
            except KeyboardInterrupt:
                is_cancelled = True
            except Exception as e:
                err_msg = str(e)

            def _finalize():
                self.btn_compile.config(state="normal", text="🚀 XUẤT FILE DOCX")
                self.status_bar.hide_progress()

                if is_cancelled:
                    self.status_bar.set_message("Đã hủy tác vụ biên dịch.")
                    messagebox.showinfo("Đã hủy", "Quá trình biên dịch đã được dừng lại.")
                elif err_msg:
                    self.status_bar.set_message(f"Lỗi: {err_msg}", is_error=True)
                    messagebox.showerror("Lỗi biên dịch DOCX", f"Không thể tạo tài liệu Word:\n{err_msg}")
                else:
                    self.status_bar.set_message("Xuất file DOCX thành công!")
                    if not settings.get("open_word", True) and compiled_path:
                        messagebox.showinfo("Thành công", f"Đã tạo thành công file Word:\n{compiled_path}")

            self.root.after(0, _finalize)

        self._compiling_thread = threading.Thread(target=_worker, daemon=True)
        self._compiling_thread.start()


def main():
    root = tk.Tk()
    app = ULNFormatterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
